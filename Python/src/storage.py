from __future__ import annotations

import csv
import os
import re
from pathlib import Path

import pandas as pd


# ---------------------------------------------------------------------------
# DNC helpers (privé)
# ---------------------------------------------------------------------------

def _normalize_company(name: str) -> str:
    """
    Zet een bedrijfsnaam om naar een genormaliseerde string:
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
    Geeft alle zinvolle varianten van een bedrijfsnaam terug.

    Voorbeelden:
      "Nabuurs - supply chain solutions" → {"nabuurs", "supply chain solutions", ...}
      "Melkweg|Fritom"                   → {"melkweg", "fritom", "melkweg fritom"}
      "Core Connect; boutique fulfilment" → {"core connect", "boutique fulfilment", ...}
      "Sanbio B.V."                      → {"sanbio"}
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


# ---------------------------------------------------------------------------
# DNC: laden en checken
# ---------------------------------------------------------------------------

def load_do_not_contact(path: str) -> set[str]:
    """
    Laad het 'Niet Benaderen' Excel-bestand en retourneer een set van
    alle genormaliseerde bedrijfsnaam-varianten.

    Gooit een FileNotFoundError als het bestand niet bestaat — het script
    mag nooit draaien zonder een geldige DNC-lijst.
    """
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"[KRITIEK] Niet-benaderen bestand niet gevonden: {path}\n"
            "Zet DNC_PATH correct in je .env of geef het juiste pad op."
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


def is_do_not_contact(company: str, dnc_set: set[str]) -> tuple[bool, str]:
    """
    Controleer of een bedrijfsnaam op de niet-benaderen lijst staat.

    Twee lagen:
      1. Exacte match op alle genormaliseerde varianten van de lead
      2. Substring-match: DNC-variant (≥5 tekens) zit in de lead of andersom

    Returns:
        (True, matched_variant) als geblokkeerd, anders (False, "")
    """
    if not company or not company.strip():
        return False, ""

    lead_variants = _extract_variants(company)

    # Laag 1: exacte match
    for v in lead_variants:
        if v in dnc_set:
            return True, v

    # Laag 2: substring-match
    # Stopwoorden die te generiek zijn om alleen een match te triggeren
    _STOPWORDS = {
        "groep", "group", "inter", "global", "solutions", "services",
        "management", "consulting", "holding", "international", "nederland",
        "netherlands", "europe", "digital", "partners", "innovations",
        "systems", "logistics", "supply", "chain", "media", "tech",
    }

    norm_lead = _normalize_company(company)
    for dnc_entry in dnc_set:
        # Minimaal 8 tekens én geen stopwoord
        if len(dnc_entry) >= 8 and dnc_entry not in _STOPWORDS:
            if dnc_entry in norm_lead or norm_lead in dnc_entry:
                return True, dnc_entry

    return False, ""


# ---------------------------------------------------------------------------
# Suppressie (al verzonden / afgemeld e-mailadressen)
# ---------------------------------------------------------------------------

def append_suppression(path: str, email: str) -> None:
    """Voeg een email toe aan de suppressielijst zodat hij nooit nogmaals verstuurd wordt."""
    log_path = Path(path)
    log_path.parent.mkdir(parents=True, exist_ok=True)
    
    write_header = not log_path.exists() or log_path.stat().st_size == 0
    
    with open(log_path, 'a', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=["email"])
        if write_header:
            writer.writeheader()
        writer.writerow({"email": email})

def load_suppression(path: str) -> set[str]:
    """
    Laad de suppressielijst (CSV met een kolom 'email').
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

    print(f"[SUPPRESSION] {len(suppressed)} e-mailadressen geladen uit {path}")
    return suppressed


# ---------------------------------------------------------------------------
# Send log
# ---------------------------------------------------------------------------

_SEND_LOG_FIELDS = ["email", "company", "title", "status", "message_id", "error", "subject", "body"]


def append_send_log(path: str, record: dict) -> None:
    """
    Voeg één record toe aan het send-log CSV-bestand.
    Maakt het bestand (inclusief header) aan als het nog niet bestaat.
    """
    log_path = Path(path)
    log_path.parent.mkdir(parents=True, exist_ok=True)

    write_header = not log_path.exists() or log_path.stat().st_size == 0

    with open(log_path, 'a', newline='', encoding='utf-8') as f:
        writer = csv.DictWriter(f, fieldnames=_SEND_LOG_FIELDS, extrasaction='ignore')
        if write_header:
            writer.writeheader()
        writer.writerow(record)