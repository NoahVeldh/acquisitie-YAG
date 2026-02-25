"""
main.py — YAG Acquisitie Tool — CLI

Gebruik:
    python main.py

Vereisten:
    - consultants/<naam>.env per consultant (zie .env.example)
    - credentials/service_account.json (voor Google Sheets)
    - data/Niet Benaderen.xlsx (DNC lijst)

Flow:
    1. Kies consultant → laadt automatisch het juiste .env bestand
    2. Leads ophalen via Lusha
    3. Leads enrichen (email / telefoon / LinkedIn)
    4. AI berichten genereren → terugschrijven naar Sheet
    5. Mails versturen (dry-run of echt)
"""

from __future__ import annotations

import os
import sys
import time
from pathlib import Path

from dotenv import load_dotenv, dotenv_values


# ══════════════════════════════════════════════════════════════════════════
# Consultant selectie — wordt uitgevoerd vóór alles
# ══════════════════════════════════════════════════════════════════════════

CONSULTANTS_DIR = Path("consultants")


def _load_consultant_env() -> str:
    CONSULTANTS_DIR.mkdir(exist_ok=True)
    env_files = sorted(CONSULTANTS_DIR.glob("*.env"))

    print("═" * 50)
    print("  YAG Acquisitie Tool")
    print("═" * 50)
    print()

    if not env_files:
        print("  ⚠ Geen consultant profielen gevonden in consultants/")
        print(f"  Maak een bestand aan via: cp .env.example consultants/jounaam.env")
        print(f"  En vul het in met jouw gegevens.\n")
        sys.exit(1)

    profiles: list[tuple[Path, str]] = []
    for env_file in env_files:
        values = dotenv_values(env_file)
        display_name = values.get("SENDER_NAME", env_file.stem.capitalize())
        vestiging    = values.get("VESTIGING_DEFAULT", "")
        label = f"{display_name}" + (f"  ({vestiging})" if vestiging else "")
        profiles.append((env_file, label))

    print("  Wie ben je?\n")
    for i, (_, label) in enumerate(profiles, 1):
        print(f"    [{i}] {label}")
    print(f"    [n] Nieuw profiel aanmaken")
    print()

    while True:
        raw = input("  > ").strip().lower()

        if raw == "n":
            _create_new_profile()
            print("\n  Profiel aangemaakt. Herstart het script.\n")
            sys.exit(0)

        if raw.isdigit() and 1 <= int(raw) <= len(profiles):
            chosen_file, chosen_label = profiles[int(raw) - 1]
            load_dotenv(chosen_file, override=True)
            print(f"\n  ✅ Ingelogd als: {chosen_label}\n")
            return chosen_label

        print(f"  Ongeldige keuze. Vul een getal in van 1 t/m {len(profiles)} of 'n'.")


def _create_new_profile() -> None:
    print("\n  ── Nieuw profiel aanmaken ──────────────────────")

    voornaam     = input("  Voornaam (voor bestandsnaam, geen spaties): ").strip().lower()
    sender_name  = input("  Volledige naam (bijv. Rick op het Veld): ").strip()
    sender_email = input("  E-mailadres: ").strip()
    sender_phone = input("  Telefoonnummer: ").strip()
    studie       = input("  Studie [Technische Bedrijfskunde]: ").strip() or "Technische Bedrijfskunde"
    universiteit = input("  Universiteit [TU Eindhoven]: ").strip() or "TU Eindhoven"
    vestiging    = input("  Vestiging (bijv. Eindhoven, Tilburg): ").strip()

    if not voornaam or not sender_email:
        print("  ❌ Voornaam en e-mailadres zijn verplicht.")
        return

    example_path = Path(".env.example")
    if example_path.exists():
        template = example_path.read_text(encoding="utf-8")

    filled = (
        template
        .replace("Rick op het Veld",                    sender_name)
        .replace("rick.ophetveld@youngadvisorygroup.nl", sender_email)
        .replace("+31 6 42 48 16 27",                   sender_phone)
        .replace("Technische Bedrijfskunde",             studie)
        .replace("TU Eindhoven",                         universiteit)
    )

    if vestiging and "VESTIGING_DEFAULT" not in filled:
        filled += f"\nVESTIGING_DEFAULT={vestiging}\n"

    token_path = f"credentials/token_{voornaam}.json"
    filled = filled.replace(
        "TOKEN_JSON=credentials/token.json",
        f"TOKEN_JSON={token_path}",
    )

    output_path = CONSULTANTS_DIR / f"{voornaam}.env"
    output_path.write_text(filled, encoding="utf-8")
    print(f"\n  ✅ Profiel opgeslagen: {output_path}")
    print(f"  ℹ  Token pad: {token_path}  (wordt aangemaakt bij eerste Gmail login)")


# ── Laad consultant profiel VOOR alle andere imports ──────────────────────
_active_consultant = _load_consultant_env()

from src.config import Col, AIStatus, MailStatus
from src.sheets import (
    get_sheets_client, open_sheet, ensure_header,
    get_all_rows, get_rows_for_ai, get_rows_for_mail,
    append_lusha_contacts, update_enriched_contact,
    set_ai_status, set_ai_result, set_ai_tokens, set_ai_error,
    set_mail_status, get_existing_contact_ids,
    cleanup_sheet,
    append_send_log_sheet, load_suppressed_emails, load_contacted_companies,
)
from src.lusha import LushaClient, ICP_PRESETS
from src.ai_gen import AIGenerator
from src.storage import load_do_not_contact, is_do_not_contact
from src.gmail_send import send_email, verify_connection


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


# ── Lusha paginateller ────────────────────────────────────────────────────

_PAGE_STATE_PATH = Path("output/lusha_page_state.json")


def _load_page_state() -> dict:
    _PAGE_STATE_PATH.parent.mkdir(exist_ok=True)
    if _PAGE_STATE_PATH.exists():
        try:
            import json
            return json.loads(_PAGE_STATE_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def _save_page_state(state: dict) -> None:
    import json
    _PAGE_STATE_PATH.parent.mkdir(exist_ok=True)
    _PAGE_STATE_PATH.write_text(json.dumps(state, indent=2), encoding="utf-8")
    _separator()


def _confirm(prompt: str = "Doorgaan? (j/n): ") -> bool:
    return input(prompt).strip().lower() in {"j", "ja", "y", "yes"}


def _pick(options: list[str], prompt: str = "Kies een optie: ") -> int:
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
    config = {
        "spreadsheet_id":      _env("SPREADSHEET_ID"),
        "worksheet_name":      _env("WORKSHEET_NAME", "Sheet1"),
        "service_account":     _env("SERVICE_ACCOUNT_JSON", "credentials/service_account.json"),
        "gmail_app_password":  _env("GMAIL_APP_PASSWORD"),
        "lusha_api_key":       _env("LUSHA_API_KEY"),
        "openai_api_key":      _env("OPENAI_API_KEY"),
        "sender_name":         _env("SENDER_NAME"),
        "sender_email":        _env("SENDER_EMAIL"),
        "sender_phone":        _env("SENDER_PHONE"),
        "sender_linkedin":     _env("SENDER_LINKEDIN", ""),
        "studie":              _env("STUDIE", "Technische Bedrijfskunde"),
        "universiteit":        _env("UNIVERSITEIT", "TU Eindhoven"),
        "subject_template":    _env("SUBJECT_TEMPLATE", "Young Advisory Group x {company}"),
        "use_web_search":      _env_bool("USE_WEB_SEARCH", True),
        "dry_run":             _env_bool("DRY_RUN", True),
        "max_emails":          _env_int("MAX_EMAILS", 20),
        "rate_limit_sec":      float(_env("RATE_LIMIT_SEC", "2")),
        "dnc_path":            _env("DNC_PATH", "data/Niet Benaderen.xlsx"),
        # Pad naar logo bestand (PNG of JPG). Leeg = geen logo in handtekening.
        # Aanbevolen: assets/logo.png  (lossless, transparantie, beste compatibiliteit)
        "logo_path":           _env("LOGO_PATH", ""),
        # Leeg = geen bijlage. Bestand niet gevonden = waarschuwing + mail
        # verstuurd zonder bijlage.
        "attachment_pdf":      _env("ATTACHMENT_PDF", ""),
        "vestiging_default":   _env("VESTIGING_DEFAULT", "Eindhoven-Tilburg"),
        "type_default":        _env("TYPE_DEFAULT", "Cold"),
        "gevallen_default":    _env("GEVALLEN_DEFAULT", ""),
        "hoe_contact_default": _env("HOE_CONTACT_DEFAULT", "Lusha"),
        "industry_ids_default": [
            int(x.strip()) for x in _env("INDUSTRY_IDS_DEFAULT", "").split(",")
            if x.strip().isdigit()
        ],
    }

    errors = []
    if not config["spreadsheet_id"]:
        errors.append("SPREADSHEET_ID ontbreekt in .env")
    if not config["sender_name"]:
        errors.append("SENDER_NAME ontbreekt in .env")
    if not config["sender_email"]:
        errors.append("SENDER_EMAIL ontbreekt in .env")
    if not config["gmail_app_password"] and not config["dry_run"]:
        errors.append("GMAIL_APP_PASSWORD ontbreekt — vereist voor echte verzending")

    if errors:
        print("\n[CONFIG] ❌ Configuratie onvolledig:")
        for e in errors:
            print(f"  • {e}")
        print("\nVul consultants/<naam>.env in op basis van consultants/.env.example en herstart.")
        sys.exit(1)

    return config


# ══════════════════════════════════════════════════════════════════════════
# Stap 1 — Leads ophalen via Lusha
# ══════════════════════════════════════════════════════════════════════════

def step_lusha_search(cfg: dict, sheet) -> None:
    _header("STAP 1 — Leads ophalen via Lusha")

    lusha = LushaClient(cfg["lusha_api_key"])

    print("\nWelk ICP profiel wil je gebruiken?")
    preset_keys = list(ICP_PRESETS.keys()) + ["Eigen filters"]
    choice = _pick(preset_keys)

    if choice < len(ICP_PRESETS):
        preset_name = preset_keys[choice]
        filters = dict(ICP_PRESETS[preset_name])
        print(f"\n  Preset: {preset_name}")
    else:
        preset_name = "eigen"
        print("\n  (Voer je eigen filters in)")
        filters = {
            "countries":     [input("  Land (bijv. Netherlands): ").strip() or "Netherlands"],
            "company_sizes": [{"min": int(input("  Min medewerkers: ") or 51),
                               "max": int(input("  Max medewerkers: ") or 1000)}],
            "industry_ids":  [],
            "job_titles":    [t.strip() for t in input("  Functietitels (kommagescheiden): ").split(",")],
        }

    default_industry_ids = cfg.get("industry_ids_default", [])
    current_ids = filters.get("industry_ids") or default_industry_ids

    def _ids_to_label(ids, industry_list):
        if not ids:
            return "Alle industrieën"
        names = [ind["main_industry"] for ind in industry_list if ind["main_industry_id"] in ids]
        return ", ".join(names) if names else str(ids)

    print(f"\n  Industrie: ", end="")
    industry_list = []
    try:
        industry_list = lusha.get_industries()
        print(_ids_to_label(current_ids, industry_list))
    except Exception:
        print(str(current_ids) if current_ids else "Alle (kon niet ophalen)")

    if _confirm("  Industrie wijzigen? (j/n): "):
        if not industry_list:
            print("  [LUSHA] Kon industrieënlijst niet ophalen.")
        else:
            print()
            for ind in industry_list:
                print(f"    [{ind['main_industry_id']:>3}] {ind['main_industry']}")
            raw = input("\n  Geef één of meer IDs op (kommagescheiden, Enter = alle): ").strip()
            if raw:
                current_ids = [int(x.strip()) for x in raw.split(",") if x.strip().isdigit()]
                print(f"  ✅ Industrie ingesteld: {_ids_to_label(current_ids, industry_list)}")
            else:
                current_ids = []
                print("  ✅ Alle industrieën")

    filters["industry_ids"] = current_ids

    preset_key = preset_name if choice < len(ICP_PRESETS) else "eigen"
    page_state = _load_page_state()
    start_page = page_state.get(preset_key, 1)

    override = input(f"\n  📄 Lusha pagina: {start_page}  (Enter = overnemen, of typ ander getal): ").strip()
    if override.isdigit():
        start_page = int(override)

    saved_meta  = page_state.get("_meta", {})
    consultant  = saved_meta.get("consultant",  cfg["sender_name"])
    vestiging   = saved_meta.get("vestiging",   cfg.get("vestiging_default", "Eindhoven-Tilburg"))
    type_       = saved_meta.get("type_",       cfg.get("type_default", "Koud contact"))
    gevallen    = saved_meta.get("gevallen",    cfg.get("gevallen_default", ""))
    hoe_contact = saved_meta.get("hoe_contact", cfg.get("hoe_contact_default", "Lusha"))

    print(f"\n  Meta-velden: {consultant} | {vestiging} | {type_} | {hoe_contact}")
    if _confirm("  Wijzigen? (j/n): "):
        consultant  = input(f"  Consultant      [{consultant}]: ").strip() or consultant
        vestiging   = input(f"  Vestiging       [{vestiging}]: ").strip() or vestiging
        type_       = input(f"  Type            [{type_}]: ").strip() or type_
        gevallen    = input(f"  Gevallen/sector [{gevallen or 'leeg'}]: ").strip() or gevallen
        hoe_contact = input(f"  Hoe contact     [{hoe_contact}]: ").strip() or hoe_contact

    page_state["_meta"] = {
        "consultant": consultant, "vestiging": vestiging,
        "type_": type_, "gevallen": gevallen, "hoe_contact": hoe_contact,
    }

    print(f"\n[LUSHA] Ophalen: 1 pagina vanaf pagina {start_page}...")
    contacts, request_id = lusha.search_multiple_pages(
        num_pages=1,
        start_page=start_page,
        **filters,
    )

    if not contacts:
        print("[LUSHA] Geen contacten gevonden.")
        return

    existing_ids = get_existing_contact_ids(sheet)
    new_contacts = [c for c in contacts if str(c.get("contactId", "")) not in existing_ids]
    skipped = len(contacts) - len(new_contacts)

    print(f"\n[LUSHA] {len(contacts)} gevonden, {skipped} al in sheet, {len(new_contacts)} nieuw.")

    if not new_contacts:
        print("[LUSHA] Niets toe te voegen.")
        return

    if input(f"Voeg {len(new_contacts)} leads toe aan de sheet? (Enter = ja, n = nee): ").strip().lower() in ("n", "nee", "no"):
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

    page_state[preset_key] = start_page + 1
    _save_page_state(page_state)

    print("\n[DNC] DNC-lijst controleren op nieuwe leads...")
    try:
        dnc_set = load_do_not_contact(cfg["dnc_path"])
    except FileNotFoundError as e:
        print(f"[DNC] ⚠ {e}")
        print("[DNC] DNC-check overgeslagen — zet DNC_PATH correct in .env")
        return

    all_rows   = get_all_rows(sheet)
    dnc_marked = 0

    for row in all_rows:
        company     = row[Col.COMPANY]
        mail_status = row[Col.MAIL_STATUS]

        if mail_status not in ("", "PENDING"):
            continue

        blocked, matched = is_do_not_contact(company, dnc_set)
        if blocked:
            set_mail_status(sheet, row["row_number"], MailStatus.DNC)
            print(f"  🚫 DNC: {company}  (match: '{matched}')")
            dnc_marked += 1

    if dnc_marked:
        print(f"[DNC] {dnc_marked} lead(s) gemarkeerd als 🚫 DNC.")
    else:
        print("[DNC] ✅ Geen matches gevonden.")


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
        and row[Col.MAIL_STATUS] not in (MailStatus.DNC, MailStatus.SUPPRESSED)
        and row[Col.CONTACT_ID]
        and row[Col.REQUEST_ID]
    ]

    if not to_enrich:
        print("[ENRICH] Geen rijen gevonden die verrijkt moeten worden.")
        return

    print(f"[ENRICH] {len(to_enrich)} leads te enrichen.")

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

        enriched_map = {e["contact_id"]: e for e in enriched}

        for row in rows:
            cid  = row[Col.CONTACT_ID]
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

    dry_run_rows = [r for r in rows if r[Col.AI_STATUS] == AIStatus.DRY_RUN]

    print(f"[AI] {len(rows)} leads klaar voor AI generatie.")
    if dry_run_rows:
        print(f"     waarvan {len(dry_run_rows)} 🔴 DRY RUN (preview) die opnieuw aangeboden worden")
    print()

    for i, row in enumerate(rows[:10], 1):
        status_tag = " 🔴" if row[Col.AI_STATUS] == AIStatus.DRY_RUN else ""
        print(f"  {i:>3}. {row[Col.COMPANY]:<30} {row[Col.FIRST_NAME]} {row[Col.LAST_NAME]}{status_tag}")
    if len(rows) > 10:
        print(f"       ... en {len(rows) - 10} meer")

    max_gen = input(f"\nHoeveel berichten genereren? (max {len(rows)}, Enter = alle): ").strip()
    limit = int(max_gen) if max_gen.isdigit() else len(rows)
    rows = rows[:limit]

    dry_run_ai = _confirm("Dry-run (geen echte OpenAI API calls, status wordt 🔴 DRY RUN)? (j/n): ")

    ai = None
    if not dry_run_ai:
        ai = AIGenerator(
            api_key=cfg["openai_api_key"],
            sender_name=cfg["sender_name"],
            sender_email=cfg["sender_email"],
            sender_phone=cfg["sender_phone"],
            sender_linkedin=cfg["sender_linkedin"],
            studie=cfg["studie"],
            universiteit=cfg["universiteit"],
        )

    print()
    done = 0
    errors = 0
    total_tokens = 0

    for row in rows:
        name    = f"{row[Col.FIRST_NAME]} {row[Col.LAST_NAME]}".strip()
        company = row[Col.COMPANY]
        label   = f"{company} | {name}"

        set_ai_status(sheet, row["row_number"], AIStatus.RUNNING)

        try:
            if dry_run_ai:
                bericht = (
                    f"[DRY RUN PREVIEW]\n\nBeste {row[Col.FIRST_NAME]},\n\n"
                    f"[AI CONNECTIEZINNEN VOOR {company}]\n\n"
                    "... rest van de mail ..."
                )
                tokens = 0
            else:
                bericht, tokens = ai.generate(
                    first_name=row[Col.FIRST_NAME],
                    job_title=row[Col.JOB_TITLE],
                    company_name=company,
                    # LinkedIn URL ingevuld door Lusha enrich (kolom G).
                    # ai_gen.py gebruikt dit als zoekhint om via web search
                    # de website en context van het bedrijf op te halen.
                    website=row[Col.LINKEDIN_URL],
                    vestiging=row[Col.VESTIGING],
                )

            set_ai_result(sheet, row["row_number"], bericht, dry_run=dry_run_ai)
            set_ai_tokens(sheet, row["row_number"], tokens)

            total_tokens += tokens
            done += 1

            status_label = "🔴 DRY RUN" if dry_run_ai else "✅"
            token_info   = f"  ({tokens} tokens)" if tokens > 0 else ""
            print(f"  {status_label} {label}{token_info}")

        except Exception as e:
            set_ai_error(sheet, row["row_number"], str(e))
            errors += 1
            print(f"  ❌ {label} — {e}")

        time.sleep(0.3)

    print(f"\n[AI] Klaar. ✅ {done} gegenereerd, ❌ {errors} fouten.")
    if dry_run_ai:
        print(f"[AI] Status in sheet: 🔴 DRY RUN — kies optie 3 opnieuw zonder dry-run om echt te genereren.")
    if total_tokens > 0:
        print(f"[AI] Totaal verbruikt: {total_tokens} tokens deze sessie.")


# ══════════════════════════════════════════════════════════════════════════
# Stap 4 — Mails versturen
# ══════════════════════════════════════════════════════════════════════════

def step_send_mail(cfg: dict, sheet) -> None:
    _header("STAP 4 — Mails versturen")

    rows = get_rows_for_mail(sheet)

    if not rows:
        print("[MAIL] Geen leads gevonden klaar voor verzending.")
        print("       Zorg dat AI Status = ✅ DONE en Mail Status leeg/PENDING is.")
        print("       Rijen met 🔴 DRY RUN worden niet verstuurd — genereer ze eerst echt via stap 3.")
        return

    dnc_set             = load_do_not_contact(cfg["dnc_path"])
    suppressed          = load_suppressed_emails(sheet)
    contacted_companies = load_contacted_companies(sheet)

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
    if skip_dnc:
        print(f"       🚫 DNC (nieuw):       {skip_dnc}")
    print(f"       ⏭  Al gemaild:        {skip_sup}")
    print(f"       ⏭  Bedrijf al gehad:  {skip_company}")
    print(f"       ✉  Klaar voor verzend: {len(sendable)}")

    if not sendable:
        print("\n[MAIL] Niets te versturen.")
        return

    dry_run = cfg["dry_run"]
    print(f"\n  DRY_RUN = {dry_run}  (wijzig in .env of toggle hieronder)")
    if _confirm("Wil je DRY_RUN omzetten? (j/n): "):
        dry_run = not dry_run
        print(f"  DRY_RUN is nu: {dry_run}")

    max_send = min(cfg["max_emails"], len(sendable))
    max_input = input(f"\nHoeveel mails versturen? (max {max_send}, Enter = {max_send}): ").strip()
    max_send = int(max_input) if max_input.isdigit() else max_send
    sendable = sendable[:max_send]

    # ── Preview eerste mail ───────────────────────────────────────────────
    print(f"\n── Preview eerste mail ──────────────────────────────")
    first = sendable[0]
    print(f"  Aan:       {first[Col.EMAIL]}")
    print(f"  Bedrijf:   {first[Col.COMPANY]}")
    print(f"  Onderwerp: {AIGenerator.subject(first[Col.COMPANY], cfg['subject_template'])}")
    if cfg["logo_path"]:
        logo_exists = "✅" if Path(cfg["logo_path"]).exists() else "⚠ niet gevonden"
        print(f"  Logo:      {cfg['logo_path']}  {logo_exists}")
    if cfg["attachment_pdf"]:
        pdf_path = Path(cfg["attachment_pdf"])
        pdf_label = pdf_path.name if pdf_path.exists() else f"{cfg['attachment_pdf']}  ⚠ niet gevonden"
        print(f"  Bijlage:   📎 {pdf_label}")
    print(f"  Body preview:\n")
    for line in first[Col.AI_BERICHT][:400].split("\n"):
        print(f"    {line}")
    print(f"    ...")
    print(f"─────────────────────────────────────────────────────")

    if not _confirm(f"\n{'[DRY RUN] ' if dry_run else ''}Verstuur {len(sendable)} mail(s)? (j/n): "):
        print("[MAIL] Geannuleerd.")
        return

    if not dry_run:
        print("[MAIL] SMTP verbinding testen...")
        if not verify_connection(cfg["sender_email"], cfg["gmail_app_password"]):
            print("[MAIL] ❌ SMTP verbinding mislukt. Controleer SENDER_EMAIL en GMAIL_APP_PASSWORD.")
            return
        print("[MAIL] ✅ SMTP verbinding OK.")

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
            append_send_log_sheet(sheet.spreadsheet, {**log_base, "status": "DRY RUN", "message_id": "", "error": ""})
            sent += 1
        else:
            try:
                message_id = send_email(
                    sender_email=cfg["sender_email"],
                    sender_name=cfg["sender_name"],
                    app_password=cfg["gmail_app_password"],
                    to_addr=email,
                    subject=subject,
                    body_text=body,
                    sender_phone=cfg["sender_phone"],
                    sender_linkedin=cfg["sender_linkedin"],
                    vestiging=row[Col.VESTIGING] or cfg["vestiging_default"],
                    attachment_pdf=cfg["attachment_pdf"],
                    logo_path=cfg["logo_path"],
                )
                set_mail_status(sheet, row["row_number"], MailStatus.SENT, message_id=message_id)
                append_send_log_sheet(sheet.spreadsheet, {
                    **log_base, "status": "SENT", "message_id": message_id, "error": ""
                })
                print(f"  ✅ SENT → {email} | {company}")
                sent += 1
                time.sleep(cfg["rate_limit_sec"])

            except Exception as e:
                err_str = str(e)
                set_mail_status(sheet, row["row_number"], MailStatus.ERROR, error=err_str)
                append_send_log_sheet(sheet.spreadsheet, {
                    **log_base, "status": "ERROR", "message_id": "", "error": err_str
                })
                print(f"  ❌ ERROR → {email} | {e}")
                errors += 1

    print(f"\n[MAIL] Klaar. ✅ {sent} verstuurd, ❌ {errors} fouten.")
    if errors:
        print(f"       Zie het 'Send Log' tabblad in de spreadsheet voor details.")


# ══════════════════════════════════════════════════════════════════════════
# Stap 5 — Overzicht
# ══════════════════════════════════════════════════════════════════════════

def step_overview(cfg: dict, sheet) -> None:
    _header("OVERZICHT — Huidige staat van de sheet")

    rows = get_all_rows(sheet)

    if not rows:
        print("[OVERZICHT] Sheet is leeg.")
        return

    ai_statuses   = {}
    mail_statuses = {}
    enriched      = {"✅ Yes": 0, "No": 0, "overig": 0}
    total_tokens  = 0

    for row in rows:
        if not any(row[c] for c in [Col.COMPANY, Col.EMAIL]):
            continue

        ai  = row[Col.AI_STATUS]   or "PENDING"
        ml  = row[Col.MAIL_STATUS] or "PENDING"
        en  = row[Col.ENRICHED]

        ai_statuses[ai]   = ai_statuses.get(ai, 0) + 1
        mail_statuses[ml] = mail_statuses.get(ml, 0) + 1
        enriched_key = en if en in enriched else "overig"
        enriched[enriched_key] = enriched.get(enriched_key, 0) + 1

        try:
            total_tokens += int(row[Col.AI_TOKENS] or 0)
        except (ValueError, TypeError):
            pass

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

    if total_tokens > 0:
        # gpt-4.1-mini: $0.40/1M input, $1.60/1M output tokens
        # Schatting: ~80% input / ~20% output voor deze use case (grote prompts, kleine output)
        # USD/EUR wisselkoers: 1 USD = €0.8492 (xe.com, 24-02-2026)
        _BLENDED_USD_PER_TOKEN = (0.80 * 0.40 + 0.20 * 1.60) / 1_000_000  # $0.64 per 1M
        _USD_TO_EUR = 0.8492
        cost_usd = total_tokens * _BLENDED_USD_PER_TOKEN
        cost_eur = cost_usd * _USD_TO_EUR
        print(f"\n  AI Tokenverbruik (cumulatief in sheet): {total_tokens:,} tokens")
        print(f"  Geschatte kosten:                       €{cost_eur:.2f}"
              f"  (${cost_usd:.4f}, blended $0.64/1M, koers €0.8492)")

    print(f"\n  Config:")
    print(f"    Consultant: {cfg['sender_name']}")
    print(f"    DRY_RUN:    {cfg['dry_run']}")
    print(f"    MAX_EMAILS: {cfg['max_emails']}")
    if cfg["logo_path"]:
        logo_exists = "✅" if Path(cfg["logo_path"]).exists() else "⚠ niet gevonden"
        print(f"    Logo:        {cfg['logo_path']}  {logo_exists}")
    if cfg["attachment_pdf"]:
        pdf_path = Path(cfg["attachment_pdf"])
        exists_label = "✅" if pdf_path.exists() else "⚠ niet gevonden"
        print(f"    PDF bijlage: {cfg['attachment_pdf']}  {exists_label}")

    # ── Lusha paginateller ────────────────────────────────────────────────
    page_state = _load_page_state()
    presets = {k: v for k, v in page_state.items() if k != "_meta"}

    if presets:
        print(f"\n  Lusha paginateller:")
        preset_list = list(presets.items())
        for i, (preset, page) in enumerate(preset_list, 1):
            print(f"    [{i}] {preset:<25} → pagina {page}")
        print(f"    [n] Niets wijzigen")
        print()

        raw = input("  Welk preset wil je aanpassen? ").strip().lower()

        if raw.isdigit() and 1 <= int(raw) <= len(preset_list):
            preset_name, current_page = preset_list[int(raw) - 1]
            new_page = input(f"  Nieuwe paginanummer [{current_page}]: ").strip()
            if new_page.isdigit() and int(new_page) >= 1:
                page_state[preset_name] = int(new_page)
                _save_page_state(page_state)
                print(f"  ✅ {preset_name} → pagina {new_page}")
            else:
                print("  Ongeldige invoer, niets gewijzigd.")
        # 'n' of enter → niets doen


# ══════════════════════════════════════════════════════════════════════════
# Stap 6 — Sheet opschonen
# ══════════════════════════════════════════════════════════════════════════

def step_cleanup(cfg: dict, sheet) -> None:
    _header("STAP 6 — Sheet opschonen")

    all_rows = get_all_rows(sheet)

    preview_dnc      = sum(1 for r in all_rows if r[Col.MAIL_STATUS] == MailStatus.DNC)
    preview_no_email = sum(1 for r in all_rows
                          if r[Col.ENRICHED] == "✅ Yes" and not r[Col.EMAIL])

    print(f"\n  Dit wordt opgeschoond:\n")
    print(f"    📦 DNC rijen verplaatst naar 'DNC Archief' tabblad:  {preview_dnc}")
    print(f"    🗑  Geen email gevonden na enrich:                   {preview_no_email}")
    print(f"    ℹ  🔴 DRY RUN rijen blijven staan (worden opnieuw aangeboden bij stap 4)")

    total = preview_dnc + preview_no_email
    if total == 0:
        print("\n  ✅ Sheet is al schoon, niets te doen.")
        return

    if not _confirm(f"\n  {total} rijen verwerken? (j/n): "):
        print("  Geannuleerd.")
        return

    print("\n  Bezig...")
    result = cleanup_sheet(sheet)

    print(f"\n  ✅ Klaar:")
    print(f"    📦 {result['moved_dnc']} DNC rijen → 'DNC Archief' tabblad")
    print(f"    🗑  {result['deleted_no_email']} rijen zonder email verwijderd")


# ══════════════════════════════════════════════════════════════════════════
# Main menu
# ══════════════════════════════════════════════════════════════════════════

def main() -> None:
    cfg = _load_config()

    dry_label = "🟡 DRY RUN" if cfg["dry_run"] else "🟢 LIVE"
    print(f"  {dry_label}  |  Max: {cfg['max_emails']} mails  |  Sheet: ...{cfg['spreadsheet_id'][-12:]}\n")

    print("[INIT] Verbinden met Google Sheets...")
    client = get_sheets_client(cfg["service_account"])
    sheet  = open_sheet(client, cfg["spreadsheet_id"], cfg["worksheet_name"])
    ensure_header(sheet)
    print("[INIT] ✅ Verbonden.\n")

    MENU = [
        ("1", "📥  Leads ophalen via Lusha",            lambda: step_lusha_search(cfg, sheet)),
        ("2", "🔍  Leads enrichen (email/tel/LinkedIn)", lambda: step_lusha_enrich(cfg, sheet)),
        ("3", "🤖  AI berichten genereren",              lambda: step_ai_generate(cfg, sheet)),
        ("4", "✉   Mails versturen",                     lambda: step_send_mail(cfg, sheet)),
        ("5", "📊  Overzicht bekijken",                  lambda: step_overview(cfg, sheet)),
        ("6", "🧹  Sheet opschonen",                     lambda: step_cleanup(cfg, sheet)),
        ("q", "🚪  Afsluiten",                           None),
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