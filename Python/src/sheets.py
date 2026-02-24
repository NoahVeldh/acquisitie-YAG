"""
sheets.py — Google Sheets connectie en alle lees/schrijf operaties

Verantwoordelijkheden:
  - Authenticeren met Google Sheets API via service account
  - Leads uitlezen (inclusief filtering op status)
  - Status terugschrijven per rij (AI status, mail status, etc.)
  - Header aanmaken als sheet leeg is
  - Nieuwe Lusha-leads appenden
"""

from __future__ import annotations

import os
from datetime import datetime
from typing import Optional

import gspread
from google.oauth2.service_account import Credentials

from src.config import Col, AIStatus, MailStatus, Enriched, DATA_START_ROW, TOTAL_COLS, REQUIRED_META_COLS, REQUIRED_META_NAMES

# ── Scopes ────────────────────────────────────────────────────────────────
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
]

# ── Header definitie (rij 1) ──────────────────────────────────────────────
HEADER_ROW = [
    "Company", "First Name", "Last Name", "Job Title", "Email", "Phone",
    "LinkedIn URL", "Enriched ✅", "AI Status", "Mail Status", "Datum Mail",
    "Follow-up datum", "Reactie ontvangen", "Opmerking",
    "---",                          # O separator
    "Consultant", "Vestiging", "Type", "Gevallen", "Hoe kom je aan dit contact",
    "---",                          # U separator
    "Request ID", "Contact ID", "isShown", "AI Bericht",
]

assert len(HEADER_ROW) == Col.TOTAL_COLS, (
    f"HEADER_ROW heeft {len(HEADER_ROW)} kolommen maar Col.TOTAL_COLS={Col.TOTAL_COLS}"
)


# ── Authenticatie ─────────────────────────────────────────────────────────

def get_sheets_client(service_account_json: str) -> gspread.Client:
    """
    Geeft een geauthenticeerde gspread client terug via een service account.

    Hoe maak je een service account aan:
      1. Google Cloud Console → IAM & Admin → Service Accounts
      2. Nieuwe service account aanmaken
      3. JSON key downloaden → sla op als credentials/service_account.json
      4. Deel je Google Sheet met het service account e-mailadres (Viewer/Editor)
    """
    if not os.path.exists(service_account_json):
        raise FileNotFoundError(
            f"Service account JSON niet gevonden: {service_account_json}\n"
            "Download hem via Google Cloud Console en zet hem op dit pad."
        )

    creds = Credentials.from_service_account_file(service_account_json, scopes=SCOPES)
    return gspread.authorize(creds)


def open_sheet(client: gspread.Client, spreadsheet_id: str, worksheet_name: str) -> gspread.Worksheet:
    """Open een specifiek worksheet op basis van spreadsheet ID en sheetnaam."""
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
    """Schrijf de header naar rij 1 als deze nog niet bestaat of verkeerd is."""
    existing = sheet.row_values(1)

    if existing == HEADER_ROW:
        return  # al correct

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
    """
    Lees alle data-rijen uit de sheet (exclusief header).
    Elke rij wordt een dict met kolomnummer → waarde.
    Voegt ook 'row_number' toe voor terugschrijven.
    """
    all_values = sheet.get_all_values()
    if len(all_values) < DATA_START_ROW:
        return []

    rows = []
    for i, row in enumerate(all_values[DATA_START_ROW - 1:], start=DATA_START_ROW):
        # Pad kortere rijen aan met lege strings
        padded = row + [""] * (Col.TOTAL_COLS - len(row))
        rows.append({
            "row_number": i,
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
        })
    return rows


def get_rows_for_ai(sheet: gspread.Worksheet) -> list[dict]:
    """
    Rijen die klaar zijn voor AI generatie:
      - Hebben een email
      - AI Status is leeg of PENDING
      - Verplichte meta-velden zijn ingevuld (Consultant, Vestiging, Type, Hoe contact)
    """
    all_rows = get_all_rows(sheet)
    eligible = []
    skipped_meta = 0

    for row in all_rows:
        if not row[Col.EMAIL] or "@" not in row[Col.EMAIL]:
            continue

        # DNC en suppressie overslaan — al vroeg gemarkeerd na search
        if row[Col.MAIL_STATUS] in (MailStatus.DNC, MailStatus.SUPPRESSED):
            continue

        ai_status = row[Col.AI_STATUS]
        if ai_status in (AIStatus.DONE, AIStatus.RUNNING):
            continue

        # Controleer verplichte meta-velden
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
    Rijen klaar voor verzending:
      - AI Status = DONE
      - AI Bericht is ingevuld
      - Mail Status is leeg of PENDING
    """
    all_rows = get_all_rows(sheet)
    return [
        row for row in all_rows
        if row[Col.AI_STATUS] == AIStatus.DONE
        and row[Col.AI_BERICHT]
        and row[Col.MAIL_STATUS] not in (MailStatus.SENT, MailStatus.DRY_RUN, MailStatus.DNC)
        and row[Col.EMAIL]
        and "@" in row[Col.EMAIL]
    ]


# ── Terugschrijven ────────────────────────────────────────────────────────

def _update_cell(sheet: gspread.Worksheet, row: int, col: int, value: str) -> None:
    """Update één cel (1-indexed row en col)."""
    sheet.update_cell(row, col, value)


def set_ai_status(sheet: gspread.Worksheet, row_number: int, status: str) -> None:
    _update_cell(sheet, row_number, Col.AI_STATUS, status)


def set_ai_result(sheet: gspread.Worksheet, row_number: int, bericht: str) -> None:
    """Schrijf het gegenereerde AI bericht en zet status op DONE."""
    # Batch update voor minimale API calls
    sheet.update(
        range_name=f"I{row_number}:Y{row_number}",
        values=[[
            AIStatus.DONE,                          # I: AI Status
            *[""] * (Col.AI_BERICHT - Col.AI_STATUS - 1),  # J t/m X leeg laten
            bericht,                                # Y: AI Bericht
        ]]
    )


def set_ai_error(sheet: gspread.Worksheet, row_number: int, error_msg: str) -> None:
    """Markeer rij als AI error, sla foutmelding op in Opmerking."""
    sheet.update(
        range_name=f"I{row_number}:N{row_number}",
        values=[[
            AIStatus.ERROR,   # I
            "",               # J Mail Status
            "",               # K Datum Mail
            "",               # L Follow-up
            "",               # M Reactie
            error_msg[:500],  # N Opmerking (truncate)
        ]]
    )


def set_mail_status(
    sheet: gspread.Worksheet,
    row_number: int,
    status: str,
    message_id: str = "",
    error: str = "",
) -> None:
    """Update Mail Status + Datum Mail na verzending."""
    datum = datetime.now().strftime("%d-%m-%Y %H:%M") if status == MailStatus.SENT else ""
    opmerking = error[:500] if error else ""

    sheet.update(
        range_name=f"J{row_number}:N{row_number}",
        values=[[status, datum, "", "", opmerking]]
    )


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
    """
    Voeg nieuwe Lusha contacten toe aan de sheet.
    Vult automatisch de meta-velden in.
    Returnt het aantal toegevoegde rijen.
    """
    if not contacts:
        return 0

    rows_to_add = []
    for c in contacts:
        full_name = c.get("name", "")
        parts = full_name.strip().split(" ", 1)
        first = parts[0] if parts else ""
        last  = parts[1] if len(parts) > 1 else ""

        row = [""] * Col.TOTAL_COLS
        row[Col.COMPANY - 1]     = c.get("companyName", "")
        row[Col.FIRST_NAME - 1]  = first
        row[Col.LAST_NAME - 1]   = last
        row[Col.JOB_TITLE - 1]   = c.get("jobTitle", "")
        row[Col.EMAIL - 1]       = ""          # wordt gevuld bij enrichment
        row[Col.PHONE - 1]       = ""
        row[Col.LINKEDIN_URL - 1]= ""
        row[Col.ENRICHED - 1]    = Enriched.NO
        row[Col.AI_STATUS - 1]   = AIStatus.PENDING
        row[Col.MAIL_STATUS - 1] = MailStatus.PENDING
        row[Col.CONSULTANT - 1]  = consultant
        row[Col.VESTIGING - 1]   = vestiging
        row[Col.TYPE - 1]        = type_
        row[Col.GEVALLEN - 1]    = gevallen
        row[Col.HOE_CONTACT - 1] = hoe_contact
        row[Col.REQUEST_ID - 1]  = request_id
        row[Col.CONTACT_ID - 1]  = str(c.get("contactId", ""))
        row[Col.IS_SHOWN - 1]    = "Yes" if c.get("isShown") else "No"

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
    """Schrijf enrichment data terug naar de juiste rij."""
    sheet.update(
        range_name=f"E{row_number}:H{row_number}",
        values=[[email, phone, linkedin, Enriched.YES]]
    )


# ── Duplicate check ───────────────────────────────────────────────────────

def get_existing_contact_ids(sheet: gspread.Worksheet) -> set[str]:
    """Geef alle Contact IDs terug die al in de sheet staan (voor duplicate-check)."""
    all_rows = get_all_rows(sheet)
    return {row[Col.CONTACT_ID] for row in all_rows if row[Col.CONTACT_ID]}


def get_existing_emails(sheet: gspread.Worksheet) -> set[str]:
    """Geef alle emails terug die al in de sheet staan."""
    all_rows = get_all_rows(sheet)
    return {row[Col.EMAIL] for row in all_rows if row[Col.EMAIL] and "@" in row[Col.EMAIL]}