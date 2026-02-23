"""
main.py — YAG Acquisitie Tool — CLI

Gebruik:
    python main.py

Vereisten:
    - .env ingevuld (zie .env.example)
    - credentials/service_account.json (voor Google Sheets)
    - credentials/credentials.json (voor Gmail OAuth)
    - data/Niet Benaderen.xlsx (DNC lijst)

Flow:
    1. Leads ophalen via Lusha
    2. Leads enrichen (email / telefoon / LinkedIn)
    3. AI berichten genereren → terugschrijven naar Sheet
    4. Mails versturen (dry-run of echt)
"""

from __future__ import annotations

import os
import sys
import time

from dotenv import load_dotenv

# ── Laad .env vroeg zodat alles beschikbaar is ────────────────────────────
load_dotenv()

from src.config import Col, AIStatus, MailStatus
from src.sheets import (
    get_sheets_client, open_sheet, ensure_header,
    get_all_rows, get_rows_for_ai, get_rows_for_mail,
    append_lusha_contacts, update_enriched_contact,
    set_ai_status, set_ai_result, set_ai_error,
    set_mail_status, get_existing_contact_ids,
)
from src.lusha import LushaClient, ICP_PRESETS
from src.ai_gen import AIGenerator
from src.storage import (
    load_do_not_contact, is_do_not_contact,
    load_suppression, append_suppression,
    append_send_log, load_contacted_companies,
)
from src.gmail_auth import get_gmail_service
from src.gmail_send import create_message, send_with_retry


# ══════════════════════════════════════════════════════════════════════════
# Helpers
# ══════════════════════════════════════════════════════════════════════════

def _env(key: str, default: str = "") -> str:
    return os.getenv(key, default).strip()

def _env_bool(key: str, default: bool = False) -> bool:
    v = os.getenv(key)
    if v is None:
        return default
    return v.strip().lower() in {"1", "true", "yes", "y", "on"}

def _env_int(key: str, default: int) -> int:
    v = os.getenv(key)
    if v is None:
        return default
    try:
        return int(v)
    except ValueError:
        return default


def _separator(char: str = "━", width: int = 50) -> None:
    print(char * width)


def _header(title: str) -> None:
    _separator()
    print(f"  {title}")
    _separator()


def _confirm(prompt: str = "Doorgaan? (j/n): ") -> bool:
    return input(prompt).strip().lower() in {"j", "ja", "y", "yes"}


def _pick(options: list[str], prompt: str = "Kies een optie: ") -> int:
    """Laat gebruiker een optie kiezen. Returnt 0-indexed keuze."""
    for i, opt in enumerate(options, 1):
        print(f"  [{i}] {opt}")
    while True:
        raw = input(prompt).strip()
        if raw.isdigit() and 1 <= int(raw) <= len(options):
            return int(raw) - 1
        print(f"  Ongeldige keuze. Vul een getal in van 1 t/m {len(options)}.")


# ══════════════════════════════════════════════════════════════════════════
# Initialisatie
# ══════════════════════════════════════════════════════════════════════════

def _load_config() -> dict:
    """Laad en valideer alle configuratie uit .env."""
    config = {
        # Sheets
        "spreadsheet_id":    _env("SPREADSHEET_ID"),
        "worksheet_name":    _env("WORKSHEET_NAME", "Sheet1"),
        "service_account":   _env("SERVICE_ACCOUNT_JSON", "credentials/service_account.json"),
        # Gmail
        "credentials_json":  _env("CREDENTIALS_JSON", "credentials/credentials.json"),
        "token_json":        _env("TOKEN_JSON", "credentials/token.json"),
        # Lusha
        "lusha_api_key":     _env("LUSHA_API_KEY"),
        # OpenAI
        "openai_api_key":    _env("OPENAI_API_KEY"),
        # Consultant
        "sender_name":       _env("SENDER_NAME"),
        "sender_email":      _env("SENDER_EMAIL"),
        "sender_phone":      _env("SENDER_PHONE"),
        "studie":            _env("STUDIE", "Technische Bedrijfskunde"),
        "universiteit":      _env("UNIVERSITEIT", "TU Eindhoven"),
        "subject_template":  _env("SUBJECT_TEMPLATE", "Young Advisory Group x {company}"),
        # Run
        "dry_run":           _env_bool("DRY_RUN", True),
        "max_emails":        _env_int("MAX_EMAILS", 20),
        "rate_limit_sec":    float(_env("RATE_LIMIT_SEC", "2")),
        # Paden
        "suppression_path":  _env("SUPPRESSION_PATH", "output/suppression.csv"),
        "send_log_path":     _env("SEND_LOG_PATH", "output/send_log.csv"),
        "dnc_path":          _env("DNC_PATH", "data/Niet Benaderen.xlsx"),
    }

    errors = []
    if not config["spreadsheet_id"]:
        errors.append("SPREADSHEET_ID ontbreekt in .env")
    if not config["sender_name"]:
        errors.append("SENDER_NAME ontbreekt in .env")
    if not config["sender_email"]:
        errors.append("SENDER_EMAIL ontbreekt in .env")

    if errors:
        print("\n[CONFIG] ❌ Configuratie onvolledig:")
        for e in errors:
            print(f"  • {e}")
        print("\nVul .env in op basis van .env.example en herstart.")
        sys.exit(1)

    return config


# ══════════════════════════════════════════════════════════════════════════
# Stap 1 — Leads ophalen via Lusha
# ══════════════════════════════════════════════════════════════════════════

def step_lusha_search(cfg: dict, sheet) -> None:
    _header("STAP 1 — Leads ophalen via Lusha")

    lusha = LushaClient(cfg["lusha_api_key"])

    # ICP kiezen
    print("\nWelk ICP profiel wil je gebruiken?")
    preset_keys = list(ICP_PRESETS.keys()) + ["Eigen filters"]
    choice = _pick(preset_keys)

    if choice < len(ICP_PRESETS):
        preset_name = preset_keys[choice]
        filters = ICP_PRESETS[preset_name]
        print(f"\n  Preset: {preset_name}")
    else:
        print("\n  (Voer je eigen filters in)")
        filters = {
            "countries":     [input("  Land (bijv. Netherlands): ").strip() or "Netherlands"],
            "company_sizes": [{"min": int(input("  Min medewerkers: ") or 51),
                               "max": int(input("  Max medewerkers: ") or 1000)}],
            "industry_ids":  [],
            "job_titles":    [t.strip() for t in input("  Functietitels (kommagescheiden): ").split(",")],
        }

    num_pages = int(input("\nHoeveel pagina's ophalen? (1 pagina = 10 leads): ") or "1")
    start_page = int(input("Startpagina: ") or "1")

    # Consultant meta
    print("\nVul de verplichte meta-velden in voor deze batch:")
    consultant = input(f"  Consultant naam [{cfg['sender_name']}]: ").strip() or cfg["sender_name"]
    vestiging  = input("  Vestiging (bijv. Eindhoven, Tilburg): ").strip()
    type_      = input("  Type (bijv. Cold, Warm, Referral): ").strip()
    gevallen   = input("  Gevallen/sector (bijv. Logistiek, Zorg): ").strip()
    hoe_contact= input("  Hoe contact (bijv. Lusha, LinkedIn): ").strip() or "Lusha"

    print(f"\n[LUSHA] Ophalen: {num_pages} pagina('s) vanaf pagina {start_page}...")
    contacts, request_id = lusha.search_multiple_pages(
        num_pages=num_pages,
        start_page=start_page,
        **filters,
    )

    if not contacts:
        print("[LUSHA] Geen contacten gevonden.")
        return

    # Duplicate check
    existing_ids = get_existing_contact_ids(sheet)
    new_contacts = [c for c in contacts if str(c.get("contactId", "")) not in existing_ids]
    skipped = len(contacts) - len(new_contacts)

    print(f"\n[LUSHA] {len(contacts)} gevonden, {skipped} al in sheet, {len(new_contacts)} nieuw.")

    if not new_contacts:
        print("[LUSHA] Niets toe te voegen.")
        return

    if not _confirm(f"Voeg {len(new_contacts)} leads toe aan de sheet? (j/n): "):
        print("[LUSHA] Geannuleerd.")
        return

    added = append_lusha_contacts(
        sheet=sheet,
        contacts=new_contacts,
        request_id=request_id,
        consultant=consultant,
        vestiging=vestiging,
        type_=type_,
        gevallen=gevallen,
        hoe_contact=hoe_contact,
    )
    print(f"[LUSHA] ✅ {added} leads toegevoegd aan de sheet.")


# ══════════════════════════════════════════════════════════════════════════
# Stap 2 — Leads enrichen
# ══════════════════════════════════════════════════════════════════════════

def step_lusha_enrich(cfg: dict, sheet) -> None:
    _header("STAP 2 — Leads enrichen (email / telefoon / LinkedIn)")

    lusha = LushaClient(cfg["lusha_api_key"])

    all_rows = get_all_rows(sheet)
    to_enrich = [
        row for row in all_rows
        if row[Col.ENRICHED] not in ("✅ Yes",)
        and row[Col.CONTACT_ID]
        and row[Col.REQUEST_ID]
    ]

    if not to_enrich:
        print("[ENRICH] Geen rijen gevonden die verrijkt moeten worden.")
        return

    print(f"[ENRICH] {len(to_enrich)} leads te enrichen.")

    # Groepeer op request_id (Lusha vereist dit)
    by_request: dict[str, list] = {}
    for row in to_enrich:
        rid = row[Col.REQUEST_ID]
        by_request.setdefault(rid, []).append(row)

    total_enriched = 0
    total_errors   = 0

    for request_id, rows in by_request.items():
        contact_ids = [row[Col.CONTACT_ID] for row in rows]
        print(f"\n[ENRICH] RequestID {request_id}: {len(contact_ids)} contacten...")

        try:
            enriched = lusha.enrich_contacts(request_id, contact_ids)
        except Exception as e:
            print(f"[ENRICH] ❌ Fout bij enrichment: {e}")
            total_errors += len(contact_ids)
            continue

        # Map op contact_id
        enriched_map = {e["contact_id"]: e for e in enriched}

        for row in rows:
            cid = row[Col.CONTACT_ID]
            data = enriched_map.get(cid)

            if not data:
                print(f"  ⚠ Contact ID {cid} niet terug in response.")
                continue

            update_enriched_contact(
                sheet=sheet,
                row_number=row["row_number"],
                email=data.get("email", ""),
                phone=data.get("phone", ""),
                linkedin=data.get("linkedin", ""),
            )
            total_enriched += 1
            print(f"  ✅ {row[Col.COMPANY]} | {row[Col.FIRST_NAME]} → {data.get('email', '(geen email)')}")
            time.sleep(0.2)

    print(f"\n[ENRICH] Klaar. ✅ {total_enriched} verrijkt, ❌ {total_errors} fouten.")


# ══════════════════════════════════════════════════════════════════════════
# Stap 3 — AI berichten genereren
# ══════════════════════════════════════════════════════════════════════════

def step_ai_generate(cfg: dict, sheet) -> None:
    _header("STAP 3 — AI berichten genereren")

    rows = get_rows_for_ai(sheet)

    if not rows:
        print("[AI] Geen leads gevonden die klaar zijn voor AI generatie.")
        print("     Controleer of de verplichte velden ingevuld zijn "
              "(Consultant, Vestiging, Type, Hoe contact).")
        return

    print(f"[AI] {len(rows)} leads klaar voor AI generatie.\n")

    # Overzicht tonen
    for i, row in enumerate(rows[:10], 1):
        print(f"  {i:>3}. {row[Col.COMPANY]:<30} {row[Col.FIRST_NAME]} {row[Col.LAST_NAME]}")
    if len(rows) > 10:
        print(f"       ... en {len(rows) - 10} meer")

    max_gen = input(f"\nHoeveel berichten genereren? (max {len(rows)}, Enter = alle): ").strip()
    limit = int(max_gen) if max_gen.isdigit() else len(rows)
    rows = rows[:limit]

    dry_run_ai = _confirm("Dry-run (geen echte OpenAI API calls)? (j/n): ")

    if not dry_run_ai:
        ai = AIGenerator(
            api_key=cfg["openai_api_key"],
            sender_name=cfg["sender_name"],
            sender_email=cfg["sender_email"],
            sender_phone=cfg["sender_phone"],
            studie=cfg["studie"],
            universiteit=cfg["universiteit"],
        )

    print()
    done = 0
    errors = 0

    for row in rows:
        name    = f"{row[Col.FIRST_NAME]} {row[Col.LAST_NAME]}".strip()
        company = row[Col.COMPANY]
        label   = f"{company} | {name}"

        # Markeer als RUNNING
        set_ai_status(sheet, row["row_number"], AIStatus.RUNNING)

        try:
            if dry_run_ai:
                bericht = ai.preview(
                    first_name=row[Col.FIRST_NAME],
                    company_name=company,
                    vestiging=row[Col.VESTIGING],
                ) if not dry_run_ai else (
                    f"[DRY RUN PREVIEW]\n\nBeste {row[Col.FIRST_NAME]},\n\n"
                    f"[AI CONNECTIEZINNEN VOOR {company}]\n\n"
                    "... rest van de mail ..."
                )
            else:
                bericht = ai.generate(
                    first_name=row[Col.FIRST_NAME],
                    job_title=row[Col.JOB_TITLE],
                    company_name=company,
                    website="",   # website niet meer in sheet, eventueel toe te voegen
                    vestiging=row[Col.VESTIGING],
                )

            set_ai_result(sheet, row["row_number"], bericht)
            done += 1
            print(f"  ✅ {label}")

        except Exception as e:
            set_ai_error(sheet, row["row_number"], str(e))
            errors += 1
            print(f"  ❌ {label} — {e}")

        time.sleep(0.3)  # kleine pauze om sheet API niet te overbelasten

    print(f"\n[AI] Klaar. ✅ {done} gegenereerd, ❌ {errors} fouten.")


# ══════════════════════════════════════════════════════════════════════════
# Stap 4 — Mails versturen
# ══════════════════════════════════════════════════════════════════════════

def step_send_mail(cfg: dict, sheet) -> None:
    _header("STAP 4 — Mails versturen")

    rows = get_rows_for_mail(sheet)

    if not rows:
        print("[MAIL] Geen leads gevonden klaar voor verzending.")
        print("       Zorg dat AI Status = ✅ DONE en Mail Status leeg/PENDING is.")
        return

    # Veiligheidslagen laden
    dnc_set             = load_do_not_contact(cfg["dnc_path"])
    suppressed          = load_suppression(cfg["suppression_path"])
    contacted_companies = load_contacted_companies(cfg["send_log_path"])

    # Filter leads
    sendable = []
    skip_dnc = skip_sup = skip_company = 0

    for row in rows:
        company = row[Col.COMPANY]
        email   = row[Col.EMAIL]

        blocked, reason = is_do_not_contact(company, dnc_set)
        if blocked:
            set_mail_status(sheet, row["row_number"], MailStatus.DNC)
            skip_dnc += 1
            continue

        if email in suppressed:
            set_mail_status(sheet, row["row_number"], MailStatus.SUPPRESSED)
            skip_sup += 1
            continue

        if company.lower() in contacted_companies:
            skip_company += 1
            continue

        sendable.append(row)

    print(f"\n[MAIL] {len(rows)} leads gevonden:")
    print(f"       🚫 DNC:               {skip_dnc}")
    print(f"       ⏭  Al gemaild:        {skip_sup}")
    print(f"       ⏭  Bedrijf al gehad:  {skip_company}")
    print(f"       ✉  Klaar voor verzend: {len(sendable)}")

    if not sendable:
        print("\n[MAIL] Niets te versturen.")
        return

    # Dry run instelling
    dry_run = cfg["dry_run"]
    print(f"\n  DRY_RUN = {dry_run}  (wijzig in .env of toggle hieronder)")
    if _confirm("Wil je DRY_RUN omzetten? (j/n): "):
        dry_run = not dry_run
        print(f"  DRY_RUN is nu: {dry_run}")

    max_send = min(cfg["max_emails"], len(sendable))
    max_input = input(f"\nHoeveel mails versturen? (max {max_send}, Enter = {max_send}): ").strip()
    max_send = int(max_input) if max_input.isdigit() else max_send
    sendable = sendable[:max_send]

    # Toon preview van eerste mail
    print(f"\n── Preview eerste mail ──────────────────────────────")
    first = sendable[0]
    print(f"  Aan:       {first[Col.EMAIL]}")
    print(f"  Bedrijf:   {first[Col.COMPANY]}")
    print(f"  Onderwerp: {AIGenerator.subject(first[Col.COMPANY], cfg['subject_template'])}")
    print(f"  Body preview:\n")
    preview = first[Col.AI_BERICHT][:400]
    for line in preview.split("\n"):
        print(f"    {line}")
    print(f"    ...")
    print(f"─────────────────────────────────────────────────────")

    if not _confirm(f"\n{'[DRY RUN] ' if dry_run else ''}Verstuur {len(sendable)} mail(s)? (j/n): "):
        print("[MAIL] Geannuleerd.")
        return

    # Gmail service (alleen als niet dry run)
    service = None
    if not dry_run:
        service = get_gmail_service(cfg["credentials_json"], cfg["token_json"])

    sent = errors = 0

    for row in sendable:
        email   = row[Col.EMAIL]
        company = row[Col.COMPANY]
        subject = AIGenerator.subject(company, cfg["subject_template"])
        body    = row[Col.AI_BERICHT]

        log_base = {
            "email":      email,
            "company":    company,
            "first_name": row[Col.FIRST_NAME],
            "job_title":  row[Col.JOB_TITLE],
            "consultant": row[Col.CONSULTANT],
            "vestiging":  row[Col.VESTIGING],
            "subject":    subject,
            "body":       body,
        }

        if dry_run:
            print(f"  [DRY RUN] → {email} | {company}")
            set_mail_status(sheet, row["row_number"], MailStatus.DRY_RUN)
            append_send_log(cfg["send_log_path"], {**log_base, "status": "DRY RUN", "message_id": "", "error": ""})
            sent += 1
        else:
            try:
                msg = create_message(
                    sender_name=cfg["sender_name"],
                    to_addr=email,
                    subject=subject,
                    body_text=body,
                )
                message_id = send_with_retry(service, "me", msg)

                append_suppression(cfg["suppression_path"], email)
                set_mail_status(sheet, row["row_number"], MailStatus.SENT, message_id=message_id)
                append_send_log(cfg["send_log_path"], {
                    **log_base, "status": "SENT", "message_id": message_id, "error": ""
                })
                print(f"  ✅ SENT → {email} | {company}")
                sent += 1
                time.sleep(cfg["rate_limit_sec"])

            except Exception as e:
                err_str = str(e)
                set_mail_status(sheet, row["row_number"], MailStatus.ERROR, error=err_str)
                append_send_log(cfg["send_log_path"], {
                    **log_base, "status": "ERROR", "message_id": "", "error": err_str
                })
                print(f"  ❌ ERROR → {email} | {e}")
                errors += 1

    print(f"\n[MAIL] Klaar. ✅ {sent} verstuurd, ❌ {errors} fouten.")
    if errors:
        print(f"       Zie {cfg['send_log_path']} voor details.")


# ══════════════════════════════════════════════════════════════════════════
# Stap 5 — Overzicht
# ══════════════════════════════════════════════════════════════════════════

def step_overview(cfg: dict, sheet) -> None:
    _header("OVERZICHT — Huidige staat van de sheet")

    rows = get_all_rows(sheet)

    if not rows:
        print("[OVERZICHT] Sheet is leeg.")
        return

    # Tel statussen
    ai_statuses   = {}
    mail_statuses = {}
    enriched      = {"✅ Yes": 0, "No": 0, "overig": 0}

    for row in rows:
        if not any(row[c] for c in [Col.COMPANY, Col.EMAIL]):
            continue  # lege rijen overslaan

        ai  = row[Col.AI_STATUS]  or "PENDING"
        ml  = row[Col.MAIL_STATUS] or "PENDING"
        en  = row[Col.ENRICHED]

        ai_statuses[ai]   = ai_statuses.get(ai, 0) + 1
        mail_statuses[ml] = mail_statuses.get(ml, 0) + 1
        enriched_key = en if en in enriched else "overig"
        enriched[enriched_key] = enriched.get(enriched_key, 0) + 1

    total = len([r for r in rows if r[Col.COMPANY]])

    print(f"\n  Totaal leads in sheet:  {total}")
    print(f"\n  Enrichment:")
    for k, v in enriched.items():
        print(f"    {k:<15} {v}")

    print(f"\n  AI Status:")
    for k, v in sorted(ai_statuses.items(), key=lambda x: x[1], reverse=True):
        print(f"    {k:<20} {v}")

    print(f"\n  Mail Status:")
    for k, v in sorted(mail_statuses.items(), key=lambda x: x[1], reverse=True):
        print(f"    {k:<20} {v}")

    print(f"\n  Config:")
    print(f"    Consultant: {cfg['sender_name']}")
    print(f"    DRY_RUN:    {cfg['dry_run']}")
    print(f"    MAX_EMAILS: {cfg['max_emails']}")


# ══════════════════════════════════════════════════════════════════════════
# Main menu
# ══════════════════════════════════════════════════════════════════════════

def main() -> None:
    _separator("═")
    print("  YAG Acquisitie Tool")
    _separator("═")

    # Config laden
    cfg = _load_config()
    print(f"\n  Consultant: {cfg['sender_name']}  ({cfg['sender_email']})")
    print(f"  DRY_RUN:    {cfg['dry_run']}")
    print(f"  Sheet:      {cfg['spreadsheet_id'][:20]}...\n")

    # Sheets connectie
    print("[INIT] Verbinden met Google Sheets...")
    client = get_sheets_client(cfg["service_account"])
    sheet  = open_sheet(client, cfg["spreadsheet_id"], cfg["worksheet_name"])
    ensure_header(sheet)
    print("[INIT] ✅ Verbonden.\n")

    # Menu loop
    MENU = [
        ("1", "📥  Leads ophalen via Lusha",           lambda: step_lusha_search(cfg, sheet)),
        ("2", "🔍  Leads enrichen (email/tel/LinkedIn)", lambda: step_lusha_enrich(cfg, sheet)),
        ("3", "🤖  AI berichten genereren",             lambda: step_ai_generate(cfg, sheet)),
        ("4", "✉   Mails versturen",                    lambda: step_send_mail(cfg, sheet)),
        ("5", "📊  Overzicht bekijken",                 lambda: step_overview(cfg, sheet)),
        ("q", "🚪  Afsluiten",                          None),
    ]

    while True:
        print()
        _separator("─")
        print("  Wat wil je doen?")
        _separator("─")
        for key, label, _ in MENU:
            print(f"  [{key}] {label}")
        _separator("─")

        choice = input("\n> ").strip().lower()

        if choice == "q":
            print("\nTot ziens! 👋\n")
            break

        action = next((fn for key, _, fn in MENU if key == choice and fn), None)

        if action is None:
            print(f"  Ongeldige keuze: '{choice}'")
            continue

        print()
        try:
            action()
        except KeyboardInterrupt:
            print("\n\n  [Onderbroken] Terug naar hoofdmenu.")
        except Exception as e:
            print(f"\n[FOUT] ❌ {e}")
            print("       Controleer je configuratie en probeer opnieuw.\n")
            import traceback
            traceback.print_exc()


if __name__ == "__main__":
    main()