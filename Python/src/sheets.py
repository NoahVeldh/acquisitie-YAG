"""
sheets.py — Google Sheets connectie en alle lees/schrijf operaties

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
WIJZIGINGEN T.O.V. VORIGE VERSIE:

  NIEUW — MailStatus.DRY_RUN is nu "🔴 DRY RUN" (was "DRY RUN")
    - get_rows_for_mail() sluit 🔴 DRY RUN niet meer uit
      → rijen worden opnieuw aangeboden bij de volgende verzendpoging
    - load_suppressed_emails() telt 🔴 DRY RUN niet meer als "al verstuurd"
      → e-mailadres blijft beschikbaar voor echte verzending
    - cleanup_sheet() verwijdert 🔴 DRY RUN (mail) rijen nog steeds als je
      dat handmatig aanvraagt via stap 6

EERDERE FIXES/TOEVOEGINGEN:
  - set_mail_status wist Follow-up (L) en Reactie (M) niet meer
  - set_ai_error overschrijft Mail Status (J) niet meer
  - set_ai_result overschrijft meta-velden niet meer
  - Kolom Z (AI Tokens) toegevoegd
  - AIStatus.DRY_RUN "🔴 DRY RUN" toegevoegd
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"""

from __future__ import annotations

import os
from datetime import datetime

import gspread
from google.oauth2.service_account import Credentials

from src.config import Col, AIStatus, MailStatus, Enriched, DATA_START_ROW, TOTAL_COLS, REQUIRED_META_COLS, REQUIRED_META_NAMES

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

HEADER_ROW = [
    "Company", "First Name", "Last Name", "Job Title", "Email", "Phone",
    "LinkedIn URL", "Enriched ✅", "AI Status", "Mail Status", "Datum Mail",
    "Follow-up datum", "Reactie ontvangen", "Opmerking",
    "---",
    "Consultant", "Vestiging", "Type", "Gevallen", "Hoe kom je aan dit contact",
    "---",
    "Request ID", "Contact ID", "isShown", "AI Bericht",
    "AI Tokens",
]

assert len(HEADER_ROW) == Col.TOTAL_COLS, (
    f"HEADER_ROW heeft {len(HEADER_ROW)} kolommen maar Col.TOTAL_COLS={Col.TOTAL_COLS}"
)


# ── Authenticatie ─────────────────────────────────────────────────────────

def get_sheets_client(service_account_json: str) -> gspread.Client:
    if not os.path.exists(service_account_json):
        raise FileNotFoundError(
            f"Service account JSON niet gevonden: {service_account_json}\n"
            "Download hem via Google Cloud Console en zet hem op dit pad."
        )
    creds = Credentials.from_service_account_file(service_account_json, scopes=SCOPES)
    return gspread.authorize(creds)


def open_sheet(client: gspread.Client, spreadsheet_id: str, worksheet_name: str) -> gspread.Worksheet:
    try:
        spreadsheet = client.open_by_key(spreadsheet_id)
    except gspread.SpreadsheetNotFound:
        raise ValueError(
            f"Spreadsheet niet gevonden: {spreadsheet_id}\n"
            "Controleer SPREADSHEET_ID in .env en zorg dat het service account toegang heeft."
        )
    try:
        return spreadsheet.worksheet(worksheet_name)
    except gspread.WorksheetNotFound:
        raise ValueError(
            f"Worksheet '{worksheet_name}' niet gevonden in spreadsheet.\n"
            f"Bestaande sheets: {[ws.title for ws in spreadsheet.worksheets()]}"
        )


# ── Header setup ──────────────────────────────────────────────────────────

def ensure_header(sheet: gspread.Worksheet) -> None:
    existing = sheet.row_values(1)
    if existing == HEADER_ROW:
        return
    if existing and existing[0] != "Company":
        print(f"[SHEETS] ⚠ Rij 1 heeft onverwachte waarden: {existing[:5]}...")
        confirm = input("Wil je de header overschrijven? (j/n): ").strip().lower()
        if confirm != "j":
            print("[SHEETS] Header niet aangepast.")
            return
    sheet.update(range_name="A1", values=[HEADER_ROW])
    print("[SHEETS] ✅ Header geschreven.")


# ── Leads lezen ───────────────────────────────────────────────────────────

def get_all_rows(sheet: gspread.Worksheet) -> list[dict]:
    all_values = sheet.get_all_values()
    if len(all_values) < DATA_START_ROW:
        return []

    rows = []
    for i, row in enumerate(all_values[DATA_START_ROW - 1:], start=DATA_START_ROW):
        padded = row + [""] * (Col.TOTAL_COLS - len(row))
        rows.append({
            "row_number":       i,
            Col.COMPANY:        padded[Col.COMPANY - 1].strip(),
            Col.FIRST_NAME:     padded[Col.FIRST_NAME - 1].strip(),
            Col.LAST_NAME:      padded[Col.LAST_NAME - 1].strip(),
            Col.JOB_TITLE:      padded[Col.JOB_TITLE - 1].strip(),
            Col.EMAIL:          padded[Col.EMAIL - 1].strip().lower(),
            Col.PHONE:          padded[Col.PHONE - 1].strip(),
            Col.LINKEDIN_URL:   padded[Col.LINKEDIN_URL - 1].strip(),
            Col.ENRICHED:       padded[Col.ENRICHED - 1].strip(),
            Col.AI_STATUS:      padded[Col.AI_STATUS - 1].strip(),
            Col.MAIL_STATUS:    padded[Col.MAIL_STATUS - 1].strip(),
            Col.DATUM_MAIL:     padded[Col.DATUM_MAIL - 1].strip(),
            Col.FOLLOWUP_DATUM: padded[Col.FOLLOWUP_DATUM - 1].strip(),
            Col.REACTIE:        padded[Col.REACTIE - 1].strip(),
            Col.OPMERKING:      padded[Col.OPMERKING - 1].strip(),
            Col.CONSULTANT:     padded[Col.CONSULTANT - 1].strip(),
            Col.VESTIGING:      padded[Col.VESTIGING - 1].strip(),
            Col.TYPE:           padded[Col.TYPE - 1].strip(),
            Col.GEVALLEN:       padded[Col.GEVALLEN - 1].strip(),
            Col.HOE_CONTACT:    padded[Col.HOE_CONTACT - 1].strip(),
            Col.REQUEST_ID:     padded[Col.REQUEST_ID - 1].strip(),
            Col.CONTACT_ID:     padded[Col.CONTACT_ID - 1].strip(),
            Col.IS_SHOWN:       padded[Col.IS_SHOWN - 1].strip(),
            Col.AI_BERICHT:     padded[Col.AI_BERICHT - 1].strip(),
            Col.AI_TOKENS:      padded[Col.AI_TOKENS - 1].strip(),
        })
    return rows


def get_rows_for_ai(sheet: gspread.Worksheet) -> list[dict]:
    """
    Rijen klaar voor AI generatie.
    🔴 DRY RUN rijen worden opnieuw aangeboden zodat je ze alsnog echt kunt genereren.
    """
    all_rows = get_all_rows(sheet)
    eligible = []
    skipped_meta = 0

    for row in all_rows:
        if not row[Col.EMAIL] or "@" not in row[Col.EMAIL]:
            continue
        if row[Col.MAIL_STATUS] in (MailStatus.DNC, MailStatus.SUPPRESSED):
            continue
        # DONE en RUNNING overslaan — DRY RUN mag opnieuw
        if row[Col.AI_STATUS] in (AIStatus.DONE, AIStatus.RUNNING):
            continue
        missing = [
            name for col, name in zip(REQUIRED_META_COLS, REQUIRED_META_NAMES)
            if not row[col]
        ]
        if missing:
            skipped_meta += 1
            continue
        eligible.append(row)

    if skipped_meta:
        print(f"[SHEETS] ⚠ {skipped_meta} rijen overgeslagen — verplichte velden ontbreken "
              f"({', '.join(REQUIRED_META_NAMES)}).")
    return eligible


def get_rows_for_mail(sheet: gspread.Worksheet) -> list[dict]:
    """
    Rijen klaar voor verzending.
    🔴 DRY RUN rijen worden opnieuw aangeboden — ze zijn nog niet echt verstuurd.
    Alleen ✅ SENT en 🚫 DNC worden uitgesloten.
    """
    all_rows = get_all_rows(sheet)
    return [
        row for row in all_rows
        if row[Col.AI_STATUS] == AIStatus.DONE
        and row[Col.AI_BERICHT]
        and row[Col.MAIL_STATUS] not in (MailStatus.SENT, MailStatus.DNC, MailStatus.SUPPRESSED)
        and row[Col.EMAIL]
        and "@" in row[Col.EMAIL]
    ]


# ── Terugschrijven ────────────────────────────────────────────────────────

def _update_cell(sheet: gspread.Worksheet, row: int, col: int, value) -> None:
    sheet.update_cell(row, col, value)


def set_ai_status(sheet: gspread.Worksheet, row_number: int, status: str) -> None:
    _update_cell(sheet, row_number, Col.AI_STATUS, status)


def set_ai_result(
    sheet: gspread.Worksheet,
    row_number: int,
    bericht: str,
    dry_run: bool = False,
) -> None:
    """
    Schrijf AI bericht (Y) en zet de juiste AI status (I).
    dry_run=True  → "🔴 DRY RUN"
    dry_run=False → "✅ DONE"
    """
    status = AIStatus.DRY_RUN if dry_run else AIStatus.DONE
    _update_cell(sheet, row_number, Col.AI_STATUS,  status)
    _update_cell(sheet, row_number, Col.AI_BERICHT, bericht)


def set_ai_tokens(sheet: gspread.Worksheet, row_number: int, tokens: int) -> None:
    if tokens > 0:
        _update_cell(sheet, row_number, Col.AI_TOKENS, tokens)


def set_ai_error(sheet: gspread.Worksheet, row_number: int, error_msg: str) -> None:
    _update_cell(sheet, row_number, Col.AI_STATUS, AIStatus.ERROR)
    _update_cell(sheet, row_number, Col.OPMERKING, error_msg[:500])


def set_mail_status(
    sheet: gspread.Worksheet,
    row_number: int,
    status: str,
    message_id: str = "",
    error: str = "",
) -> None:
    """
    Update Mail Status (J) en Datum Mail (K).
    Follow-up datum (L) en Reactie (M) worden NIET aangeraakt.
    """
    datum = datetime.now().strftime("%d-%m-%Y %H:%M") if status == MailStatus.SENT else ""
    _update_cell(sheet, row_number, Col.MAIL_STATUS, status)
    _update_cell(sheet, row_number, Col.DATUM_MAIL,  datum)
    if error:
        _update_cell(sheet, row_number, Col.OPMERKING, error[:500])


def set_enriched(sheet: gspread.Worksheet, row_number: int, enriched: bool) -> None:
    _update_cell(sheet, row_number, Col.ENRICHED, Enriched.YES if enriched else Enriched.NO)


# ── Nieuwe Lusha leads toevoegen ──────────────────────────────────────────

def append_lusha_contacts(
    sheet: gspread.Worksheet,
    contacts: list[dict],
    request_id: str,
    consultant: str,
    vestiging: str,
    type_: str,
    gevallen: str,
    hoe_contact: str,
) -> int:
    if not contacts:
        return 0

    rows_to_add = []
    for c in contacts:
        full_name = c.get("name", "")
        parts = full_name.strip().split(" ", 1)
        first = parts[0] if parts else ""
        last  = parts[1] if len(parts) > 1 else ""

        row = [""] * Col.TOTAL_COLS
        row[Col.COMPANY - 1]      = c.get("companyName", "")
        row[Col.FIRST_NAME - 1]   = first
        row[Col.LAST_NAME - 1]    = last
        row[Col.JOB_TITLE - 1]    = c.get("jobTitle", "")
        row[Col.EMAIL - 1]        = ""
        row[Col.PHONE - 1]        = ""
        row[Col.LINKEDIN_URL - 1] = ""
        row[Col.ENRICHED - 1]     = Enriched.NO
        row[Col.AI_STATUS - 1]    = AIStatus.PENDING
        row[Col.MAIL_STATUS - 1]  = MailStatus.PENDING
        row[Col.CONSULTANT - 1]   = consultant
        row[Col.VESTIGING - 1]    = vestiging
        row[Col.TYPE - 1]         = type_
        row[Col.GEVALLEN - 1]     = gevallen
        row[Col.HOE_CONTACT - 1]  = hoe_contact
        row[Col.REQUEST_ID - 1]   = request_id
        row[Col.CONTACT_ID - 1]   = str(c.get("contactId", ""))
        row[Col.IS_SHOWN - 1]     = "Yes" if c.get("isShown") else "No"
        row[Col.AI_TOKENS - 1]    = ""

        rows_to_add.append(row)

    sheet.append_rows(rows_to_add, value_input_option="USER_ENTERED")
    return len(rows_to_add)


def update_enriched_contact(
    sheet: gspread.Worksheet,
    row_number: int,
    email: str,
    phone: str,
    linkedin: str,
) -> None:
    sheet.update(
        range_name=f"E{row_number}:H{row_number}",
        values=[[email, phone, linkedin, Enriched.YES]]
    )


# ── Duplicate check ───────────────────────────────────────────────────────

def get_existing_contact_ids(sheet: gspread.Worksheet) -> set[str]:
    all_rows = get_all_rows(sheet)
    return {row[Col.CONTACT_ID] for row in all_rows if row[Col.CONTACT_ID]}


def get_existing_emails(sheet: gspread.Worksheet) -> set[str]:
    all_rows = get_all_rows(sheet)
    return {row[Col.EMAIL] for row in all_rows if row[Col.EMAIL] and "@" in row[Col.EMAIL]}


# ── Sheet opschonen ───────────────────────────────────────────────────────

DNC_ARCHIVE_SHEET = "DNC Archief"


def _get_or_create_archive(spreadsheet: gspread.Spreadsheet) -> gspread.Worksheet:
    try:
        archive = spreadsheet.worksheet(DNC_ARCHIVE_SHEET)
    except gspread.WorksheetNotFound:
        archive = spreadsheet.add_worksheet(title=DNC_ARCHIVE_SHEET, rows=1000, cols=Col.TOTAL_COLS)
        archive.append_row(HEADER_ROW, value_input_option="USER_ENTERED")
        archive.format("1:1", {"textFormat": {"bold": True}})
    return archive


def cleanup_sheet(sheet: gspread.Worksheet) -> dict:
    """
    Schoon de sheet op:
      - DNC rijen          → verplaatst naar 'DNC Archief' tabblad
      - Geen email na enrich → verwijderd
      - 🔴 DRY RUN (mail)  → blijft staan, wordt opnieuw aangeboden bij verzending
    """
    all_rows = get_all_rows(sheet)
    spreadsheet = sheet.spreadsheet

    moved_dnc = deleted_no_email = deleted_dry_run = 0
    dnc_rows = []
    delete_rows = []

    for row in all_rows:
        mail_status = row[Col.MAIL_STATUS]
        email       = row[Col.EMAIL]
        enriched    = row[Col.ENRICHED]

        if mail_status == MailStatus.DNC:
            dnc_rows.append(row)
        elif enriched == Enriched.YES and not email:
            delete_rows.append(row)

    if dnc_rows:
        archive = _get_or_create_archive(spreadsheet)
        archive_data = [[row.get(col, "") for col in range(1, Col.TOTAL_COLS + 1)] for row in dnc_rows]
        archive.append_rows(archive_data, value_input_option="USER_ENTERED")
        moved_dnc = len(dnc_rows)

    for row in delete_rows:
        deleted_no_email += 1

    all_to_delete = sorted(dnc_rows + delete_rows, key=lambda r: r["row_number"], reverse=True)

    if all_to_delete:
        requests = []
        for row in all_to_delete:
            idx = row["row_number"] - 1
            requests.append({
                "deleteDimension": {
                    "range": {
                        "sheetId":    sheet.id,
                        "dimension":  "ROWS",
                        "startIndex": idx,
                        "endIndex":   idx + 1,
                    }
                }
            })
        sheet.spreadsheet.batch_update({"requests": requests})

    return {
        "moved_dnc":        moved_dnc,
        "deleted_no_email": deleted_no_email,
    }


# ── Send Log tabblad ──────────────────────────────────────────────────────

SEND_LOG_SHEET   = "Send Log"
SEND_LOG_HEADERS = [
    "Timestamp", "Consultant", "Vestiging", "Company", "First Name",
    "Job Title", "Email", "Subject", "Status", "Message ID", "Error",
]


def _get_or_create_send_log(spreadsheet: gspread.Spreadsheet) -> gspread.Worksheet:
    try:
        return spreadsheet.worksheet(SEND_LOG_SHEET)
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(title=SEND_LOG_SHEET, rows=5000, cols=len(SEND_LOG_HEADERS))
        ws.append_row(SEND_LOG_HEADERS, value_input_option="USER_ENTERED")
        ws.format("1:1", {"textFormat": {"bold": True}})
        return ws


def append_send_log_sheet(spreadsheet: gspread.Spreadsheet, record: dict) -> None:
    ws = _get_or_create_send_log(spreadsheet)
    ws.append_row([
        record.get("timestamp", datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        record.get("consultant", ""),
        record.get("vestiging", ""),
        record.get("company", ""),
        record.get("first_name", ""),
        record.get("job_title", ""),
        record.get("email", ""),
        record.get("subject", ""),
        record.get("status", ""),
        record.get("message_id", ""),
        record.get("error", ""),
    ], value_input_option="USER_ENTERED")


def load_suppressed_emails(sheet: gspread.Worksheet) -> set[str]:
    """
    Haal al-verzonden e-mailadressen op.
    🔴 DRY RUN wordt NIET als verstuurd beschouwd — die mails zijn nooit echt gegaan.
    Alleen ✅ SENT telt als suppressed.
    """
    all_rows = get_all_rows(sheet)
    suppressed = {
        row[Col.EMAIL].strip().lower()
        for row in all_rows
        if row[Col.MAIL_STATUS] == MailStatus.SENT
        and row[Col.EMAIL]
    }
    if suppressed:
        print(f"[SUPPRESSION] {len(suppressed)} e-mailadressen al verstuurd (uit sheet).")
    return suppressed


def load_contacted_companies(sheet: gspread.Worksheet) -> set[str]:
    all_rows = get_all_rows(sheet)
    companies = {
        row[Col.COMPANY].strip().lower()
        for row in all_rows
        if row[Col.MAIL_STATUS] == MailStatus.SENT
        and row[Col.COMPANY]
    }
    if companies:
        print(f"[COMPANIES] {len(companies)} bedrijven al eerder benaderd (uit sheet).")
    return companies