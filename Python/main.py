"""
main.py — YAG Acquisitie Tool — CLI

Gebruik:
    python main.py

Vereisten:
    - consultants/<naam>.env per consultant (zie .env.example)
    - credentials/service_account.json (voor Google Sheets)
    - credentials/<naam>_credentials.json (voor Gmail OAuth, per consultant)
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
    """
    Zoek alle .env bestanden in de consultants/ map, laat de gebruiker
    kiezen en laad het juiste bestand. Geeft de naam van de consultant terug.

    Bestandsnaamconventie: consultants/<voornaam>.env
    Voorbeeld:             consultants/rick.env
    """
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

    # Lees SENDER_NAME uit elk .env bestand voor een nette weergave
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
            # Herstart zodat het nieuwe profiel in de lijst verschijnt
            print("\n  Profiel aangemaakt. Herstart het script.\n")
            sys.exit(0)

        if raw.isdigit() and 1 <= int(raw) <= len(profiles):
            chosen_file, chosen_label = profiles[int(raw) - 1]
            load_dotenv(chosen_file, override=True)
            print(f"\n  ✅ Ingelogd als: {chosen_label}\n")
            return chosen_label

        print(f"  Ongeldige keuze. Vul een getal in van 1 t/m {len(profiles)} of 'n'.")


def _create_new_profile() -> None:
    """Interactief een nieuw consultant profiel aanmaken."""
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

    # Lees het .env.example als basis
    example_path = Path(".env.example")
    if example_path.exists():
        template = example_path.read_text(encoding="utf-8")

    # Vervang de placeholders
    filled = (
        template
        .replace("Rick op het Veld",                   sender_name)
        .replace("rick.ophetveld@youngadvisorygroup.nl", sender_email)
        .replace("+31 6 42 48 16 27",                  sender_phone)
        .replace("Technische Bedrijfskunde",            studie)
        .replace("TU Eindhoven",                        universiteit)
    )

    # Voeg vestiging toe als extra variabele
    if vestiging and "VESTIGING_DEFAULT" not in filled:
        filled += f"\nVESTIGING_DEFAULT={vestiging}\n"

    # Zet token pad uniek per consultant
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
        "spreadsheet_id":     _env("SPREADSHEET_ID"),
        "worksheet_name":     _env("WORKSHEET_NAME", "Sheet1"),
        "service_account":    _env("SERVICE_ACCOUNT_JSON", "credentials/service_account.json"),
        # Gmail SMTP (App Password — geen OAuth nodig)
        "gmail_app_password": _env("GMAIL_APP_PASSWORD"),
        # Lusha
        "lusha_api_key":      _env("LUSHA_API_KEY"),
        # OpenAI
        "openai_api_key":     _env("OPENAI_API_KEY"),
        # Consultant
        "sender_name":        _env("SENDER_NAME"),
        "sender_email":       _env("SENDER_EMAIL"),
        "sender_phone":       _env("SENDER_PHONE"),
        "studie":             _env("STUDIE", "Technische Bedrijfskunde"),
        "universiteit":       _env("UNIVERSITEIT", "TU Eindhoven"),
        "subject_template":   _env("SUBJECT_TEMPLATE", "Young Advisory Group x {company}"),
        # Run
        "dry_run":            _env_bool("DRY_RUN", True),
        "max_emails":         _env_int("MAX_EMAILS", 20),
        "rate_limit_sec":     float(_env("RATE_LIMIT_SEC", "2")),
        # Paden
        "suppression_path":   _env("SUPPRESSION_PATH", "output/suppression.csv"),
        "send_log_path":      _env("SEND_LOG_PATH", "output/send_log.csv"),
        "dnc_path":           _env("DNC_PATH", "data/Niet Benaderen.xlsx"),
        # Meta-veld defaults (vooringevuld bij Lusha stap)
        "vestiging_default":   _env("VESTIGING_DEFAULT", "Eindhoven-Tilburg"),
        "type_default":        _env("TYPE_DEFAULT", "Cold"),
        "gevallen_default":    _env("GEVALLEN_DEFAULT", ""),
        "hoe_contact_default": _env("HOE_CONTACT_DEFAULT", "Lusha"),
        # Lusha industrie default (lijst van IDs, leeg = alle industrieën)
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

    # ── ICP kiezen ────────────────────────────────────────────────────────
    print("\nWelk ICP profiel wil je gebruiken?")
    preset_keys = list(ICP_PRESETS.keys()) + ["Eigen filters"]
    choice = _pick(preset_keys)

    if choice < len(ICP_PRESETS):
        preset_name = preset_keys[choice]
        filters = dict(ICP_PRESETS[preset_name])   # kopie zodat we kunnen aanpassen
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

    # ── Industrie kiezen / bevestigen ─────────────────────────────────────
    default_industry_ids = cfg.get("industry_ids_default", [])
    current_ids = filters.get("industry_ids") or default_industry_ids

    # Haal industrie-namen op voor weergave
    def _ids_to_label(ids, industry_list):
        if not ids:
            return "Alle industrieën"
        names = []
        for ind in industry_list:
            if ind["main_industry_id"] in ids:
                names.append(ind["main_industry"])
        return ", ".join(names) if names else str(ids)

    print(f"\n  Industrie: ", end="")
    industry_list = []
    try:
        industry_list = lusha.get_industries()
        current_label = _ids_to_label(current_ids, industry_list)
        print(current_label)
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

    # ── Pagina's — willekeurige startpagina om duplicaten te voorkomen ──────
    import random
    random_page = random.randint(1, 50)

    num_pages = 1
    print(f"\n  Startpagina: {random_page} (willekeurig)")
    override = input(f"  Andere startpagina? (Enter = {random_page}, of typ getal): ").strip()
    start_page = int(override) if override.isdigit() else random_page

    # ── Meta-velden ───────────────────────────────────────────────────────
    default_vestiging   = cfg.get("vestiging_default", "Eindhoven-Tilburg")
    default_type        = cfg.get("type_default", "Koud contact")
    default_gevallen    = cfg.get("gevallen_default", "")
    default_hoe_contact = cfg.get("hoe_contact_default", "Lusha")

    print(f"\nMeta-velden (Enter = standaardwaarde overnemen):")
    consultant  = input(f"  Consultant      [{cfg['sender_name']}]: ").strip() or cfg["sender_name"]
    vestiging   = input(f"  Vestiging       [{default_vestiging}]: ").strip() or default_vestiging
    type_       = input(f"  Type            [{default_type}]: ").strip() or default_type
    gevallen    = input(f"  Gevallen/sector [{default_gevallen or 'leeg'}]: ").strip() or default_gevallen
    hoe_contact = input(f"  Hoe contact     [{default_hoe_contact}]: ").strip() or default_hoe_contact

    # ── Ophalen ───────────────────────────────────────────────────────────
    print(f"\n[LUSHA] Ophalen: {num_pages} pagina(\'s) vanaf pagina {start_page}...")
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

    # ── DNC scan direct na toevoegen ──────────────────────────────────────
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

        # Alleen rijen die net zijn toegevoegd (PENDING) en nog geen status hebben
        if mail_status not in ("", "PENDING"):
            continue

        blocked, matched = is_do_not_contact(company, dnc_set)
        if blocked:
            set_mail_status(sheet, row["row_number"], MailStatus.DNC)
            print(f"  🚫 DNC: {company}  (match: '{matched}')")
            dnc_marked += 1

    if dnc_marked:
        print(f"[DNC] {dnc_marked} lead(s) gemarkeerd als 🚫 DNC — worden overgeslagen bij enrich en AI.")
    else:
        print("[DNC] ✅ Geen matches gevonden.")



# ══════════════════════════════════════════════════════════════════════════
# Stap 2 — Leads enrichen
# ══════════════════════════════════════════════════════════════════════════

def step_lusha_enrich(cfg: dict, sheet) -> None:
    _header("STAP 2 — Leads enrichen (email / telefoon / LinkedIn)")

    lusha = LushaClient(cfg["lusha_api_key"])

    all_rows = get_all_rows(sheet)
    skip_not_shown = 0
    to_enrich = []
    for row in all_rows:
        if row[Col.ENRICHED] in ("✅ Yes",):
            continue
        if row[Col.MAIL_STATUS] in (MailStatus.DNC, MailStatus.SUPPRESSED):
            continue
        if not row[Col.CONTACT_ID] or not row[Col.REQUEST_ID]:
            continue
        if row[Col.IS_SHOWN].strip().lower() != "yes":
            skip_not_shown += 1
            continue
        to_enrich.append(row)

    if skip_not_shown:
        print(f"[ENRICH] ⏭ {skip_not_shown} lead(s) overgeslagen — isShown=No (geen credits beschikbaar bij Lusha).")

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
    # DNC is al gecheckt na search — dit is een tweede controle voor het geval
    # de DNC lijst is bijgewerkt na de laatste search.
    dnc_set             = load_do_not_contact(cfg["dnc_path"])
    suppressed          = load_suppression(cfg["suppression_path"])
    contacted_companies = load_contacted_companies(cfg["send_log_path"])

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
        print(f"       🚫 DNC (nieuw):       {skip_dnc}  ← DNC lijst bijgewerkt na search")
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

    # SMTP verbinding testen vóór batch (alleen bij echte verzending)
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
            append_send_log(cfg["send_log_path"], {**log_base, "status": "DRY RUN", "message_id": "", "error": ""})
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
                )
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
    # Config laden (consultant .env al geladen door _load_consultant_env())
    cfg = _load_config()

    # Toon actieve sessie info
    dry_label = "🟡 DRY RUN" if cfg["dry_run"] else "🟢 LIVE"
    print(f"  {dry_label}  |  Max: {cfg['max_emails']} mails  |  Sheet: ...{cfg['spreadsheet_id'][-12:]}\n")

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