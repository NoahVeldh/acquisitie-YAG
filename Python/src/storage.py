"""
storage.py — Lokale opslag: DNC lijst, suppressie en send-log

De sheet is de primary source of truth voor leads en statussen.
Deze module beheert de lokale veiligheidslagen:
  - DNC (Do Not Contact): bedrijven die nooit benaderd mogen worden
  - Suppressie: e-mailadressen die al een mail ontvangen hebben
  - Send log: volledige audit trail van alle verzendpogingen
"""

from __future__ import annotations

import csv
import os
import re
from pathlib import Path

import pandas as pd


# ══════════════════════════════════════════════════════════════════════════
# DNC — Do Not Contact
# ══════════════════════════════════════════════════════════════════════════

def _normalize_company(name: str) -> str:
    """
    Normaliseer een bedrijfsnaam voor vergelijking:
    - lowercase
    - verwijder juridische vormen (B.V., NV, BV, VOF, Ltd, ...)
    - verwijder leestekens → spaties
    - collapse whitespace
    """
    name = name.lower()
    name = re.sub(
        r'\b(b\.?\s*v\.?|n\.?\s*v\.?|v\.?\s*o\.?\s*f\.?|ltd\.?|gmbh|inc\.?|llc|bv|nv|vof)\b',
        '',
        name,
    )
    name = re.sub(r'[^a-z0-9\s]', ' ', name)
    return re.sub(r'\s+', ' ', name).strip()


def _extract_variants(raw: str) -> set[str]:
    """
    Geef alle zinvolle varianten van een bedrijfsnaam terug.

    Voorbeelden:
      "Nabuurs - supply chain solutions" → {"nabuurs", "supply chain solutions", ...}
      "Melkweg|Fritom"                   → {"melkweg", "fritom", "melkweg fritom"}
    """
    variants: set[str] = set()
    variants.add(_normalize_company(raw))

    for sep in [';', '|', ' - ', ',']:
        if sep in raw:
            for part in raw.split(sep):
                norm = _normalize_company(part.strip())
                if len(norm) >= 3:
                    variants.add(norm)

    return {v for v in variants if v}


def load_do_not_contact(path: str) -> set[str]:
    """
    Laad het 'Niet Benaderen' Excel-bestand.
    Retourneert een set van alle genormaliseerde bedrijfsnaam-varianten.

    Gooit een FileNotFoundError als het bestand niet bestaat —
    het script mag NOOIT draaien zonder geldige DNC-lijst.
    """
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"[KRITIEK] DNC-bestand niet gevonden: {path}\n"
            "Zet DNC_PATH correct in .env of geef het juiste pad op."
        )

    if path.lower().endswith('.csv'):
        df = pd.read_csv(path)
        col = 'Bedrijf' if 'Bedrijf' in df.columns else df.columns[0]
    else:
        df = pd.read_excel(path, sheet_name='Niet Benaderen')
        col = 'Bedrijf'

    if col not in df.columns:
        raise ValueError(
            f"Kolom '{col}' niet gevonden in {path}. "
            f"Gevonden kolommen: {df.columns.tolist()}"
        )

    dnc_set: set[str] = set()
    for raw in df[col].dropna():
        dnc_set.update(_extract_variants(str(raw)))

    print(f"[DNC] {len(df)} bedrijven geladen → {len(dnc_set)} genormaliseerde varianten.")
    return dnc_set


_STOPWORDS = {
    "groep", "group", "inter", "global", "solutions", "services",
    "management", "consulting", "holding", "international", "nederland",
    "netherlands", "europe", "digital", "partners", "innovations",
    "systems", "logistics", "supply", "chain", "media", "tech",
}


def is_do_not_contact(company: str, dnc_set: set[str]) -> tuple[bool, str]:
    """
    Controleer of een bedrijfsnaam op de DNC-lijst staat.

    Twee lagen:
      1. Exacte match op alle genormaliseerde varianten van de lead
      2. Substring-match: DNC-variant (≥8 tekens) zit in de lead of andersom

    Returns:
        (True, matched_variant) als geblokkeerd, anders (False, "")
    """
    if not company or not company.strip():
        return False, ""

    lead_variants = _extract_variants(company)

    for v in lead_variants:
        if v in dnc_set:
            return True, v

    norm_lead = _normalize_company(company)
    for dnc_entry in dnc_set:
        if len(dnc_entry) >= 8 and dnc_entry not in _STOPWORDS:
            if dnc_entry in norm_lead or norm_lead in dnc_entry:
                return True, dnc_entry

    return False, ""


# ══════════════════════════════════════════════════════════════════════════
# Suppressie — al verzonden e-mailadressen
# ══════════════════════════════════════════════════════════════════════════

def load_suppression(path: str) -> set[str]:
    """
    Laad de suppressielijst (CSV met kolom 'email').
    Geeft een lege set terug als het bestand niet bestaat.
    """
    if not os.path.exists(path):
        return set()

    suppressed: set[str] = set()
    with open(path, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            email = (row.get("email") or "").strip().lower()
            if email:
                suppressed.add(email)

    print(f"[SUPPRESSION] {len(suppressed)} e-mailadressen geladen.")
    return suppressed


def append_suppression(path: str, email: str) -> None:
    """Voeg een email toe aan de suppressielijst."""
    log_path = Path(path)
    log_path.parent.mkdir(parents=True, exist_ok=True)

    write_header = not log_path.exists() or log_path.stat().st_size == 0

    with open(log_path, 'a', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=["email"], quoting=csv.QUOTE_ALL)
        if write_header:
            writer.writeheader()
        writer.writerow({"email": email.strip().lower()})


# ══════════════════════════════════════════════════════════════════════════
# Send log — volledige audit trail
# ══════════════════════════════════════════════════════════════════════════

_SEND_LOG_FIELDS = [
    "timestamp", "email", "company", "first_name", "job_title",
    "consultant", "vestiging", "status", "message_id", "error",
    "subject", "body",
]


def append_send_log(path: str, record: dict) -> None:
    """
    Voeg één record toe aan het send-log CSV.
    Maakt het bestand (inclusief header) aan als het nog niet bestaat.
    QUOTE_ALL zodat newlines en komma's in de body geen problemen geven.
    """
    from datetime import datetime

    log_path = Path(path)
    log_path.parent.mkdir(parents=True, exist_ok=True)

    write_header = not log_path.exists() or log_path.stat().st_size == 0

    if "timestamp" not in record:
        record["timestamp"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    with open(log_path, 'a', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(
            f,
            fieldnames=_SEND_LOG_FIELDS,
            extrasaction='ignore',
            quoting=csv.QUOTE_ALL,
        )
        if write_header:
            writer.writeheader()
        writer.writerow(record)


def load_contacted_companies(path: str) -> set[str]:
    """
    Laad alle bedrijven die al eerder benaderd zijn (status=SENT).
    Voorkomt dat collega's van hetzelfde bedrijf ook een mail krijgen.
    """
    if not os.path.exists(path):
        return set()

    companies: set[str] = set()
    with open(path, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        for row in reader:
            if (row.get("status") or "").strip() == "SENT":
                company = (row.get("company") or "").strip().lower()
                if company:
                    companies.add(company)

    print(f"[COMPANIES] {len(companies)} bedrijven al eerder benaderd.")
    return companies