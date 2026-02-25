"""
recent_contacts.py — Laatste Contactmomenten check

Leest de Excel met recente contactmomenten en bepaalt per bedrijf
of het nog in een cooldown-periode zit.

Cooldown regels:
  - Type is UITSLUITEND "Gemaild" en/of "Gebeld"  →  3 maanden cooldown
  - Type bevat iets anders (Gesprek, Afspraak, etc.) →  1 jaar cooldown

Als één rij voor een bedrijf blokkeert, wordt het hele bedrijf geblokkeerd.
"""

from __future__ import annotations

import re
from datetime import datetime, timedelta
from pathlib import Path


def _fix_date(datum: datetime, now: datetime) -> datetime:
    """
    Excel (met US-locale) slaat "06-02-2026" (6 feb, NL-formaat) op als
    MM/DD → 2 juni 2026. Openpyxl leest dat correct uit als datum in de
    toekomst. Als de dag ≤ 12, probeer dag en maand te swappen — als dat
    een datum in het verleden oplevert, was het waarschijnlijk verkeerd
    omgedraaid.
    """
    if datum <= now:
        return datum  # datum ligt al in het verleden, geen probleem
    if datum.day <= 12:
        try:
            swapped = datum.replace(day=datum.month, month=datum.day)
            if swapped <= now:
                return swapped
        except ValueError:
            pass
    return datum

try:
    import openpyxl
except ImportError:
    openpyxl = None  # type: ignore

# Kolom-indices in de Excel (0-based)
_COL_BEDRIJF = 2   # "Bedrijf"
_COL_DATUM   = 5   # "Datum"
_COL_TYPE    = 10  # "Type (aantal punten)"

# Types die als "licht contact" gelden → 3 maanden cooldown
_LIGHT_CONTACT_TYPES = {"gemaild", "gebeld", "gemailed"}

# Cooldown periodes
_COOLDOWN_LIGHT  = timedelta(days=90)   # 3 maanden
_COOLDOWN_HEAVY  = timedelta(days=365)  # 1 jaar


def _normalize_type(raw_type: str) -> set[str]:
    """
    Splits een type-string op ";" en normaliseert elk onderdeel.
    Verwijdert puntenwaarden tussen haakjes, bijv. "Gemaild (2)" → "gemaild".
    """
    parts = set()
    for part in raw_type.split(";"):
        clean = re.sub(r"\s*\(.*?\)", "", part).strip().lower()
        if clean and clean != "#ref!":
            parts.add(clean)
    return parts


def _is_light_contact(types: set[str]) -> bool:
    """True als het type uitsluitend uit licht contact bestaat (mailen/bellen)."""
    return bool(types) and types.issubset(_LIGHT_CONTACT_TYPES)


def load_recent_contacts(path: str) -> dict[str, list[tuple[datetime, str]]]:
    """
    Laad de Excel en geef een dict terug:
        { bedrijfsnaam_lower: [(datum, type_raw), ...] }

    Alleen rijen met een geldig bedrijf én datum worden meegenomen.
    """
    if openpyxl is None:
        raise ImportError(
            "openpyxl is niet geïnstalleerd. Installeer het met: pip install openpyxl"
        )

    p = Path(path)
    if not p.exists():
        raise FileNotFoundError(
            f"Contactmomenten bestand niet gevonden: {path}\n"
            "Controleer RECENT_CONTACTS_PATH in .env."
        )

    wb = openpyxl.load_workbook(p, read_only=True, data_only=True)
    ws = wb.active

    contacts: dict[str, list[tuple[datetime, str]]] = {}
    skipped = 0
    now = datetime.now()

    # Tekst-formaten om te proberen als openpyxl een string teruggeeft
    _TEXT_FORMATS = [
        "%d-%m-%Y",   # NL: 06-02-2026 → 6 feb
        "%m-%d-%Y",   # US: 02-06-2026 → 6 feb (in context: swap)
        "%d/%m/%Y",   # NL: 06/02/2026
        "%m/%d/%Y",   # US: 02/06/2026
        "%Y-%m-%d",   # ISO: 2026-02-06
    ]

    for row in ws.iter_rows(min_row=2, values_only=True):
        bedrijf  = row[_COL_BEDRIJF]
        datum    = row[_COL_DATUM]
        type_raw = row[_COL_TYPE]

        if not bedrijf or not datum or not type_raw:
            skipped += 1
            continue

        if isinstance(datum, str):
            parsed = None
            for fmt in _TEXT_FORMATS:
                try:
                    parsed = datetime.strptime(datum.strip(), fmt)
                    break
                except ValueError:
                    continue
            if parsed is None:
                skipped += 1
                continue
            datum = parsed

        if not isinstance(datum, datetime):
            skipped += 1
            continue

        # Corrigeer datums die door Excel US-locale verkeerd zijn opgeslagen
        datum = _fix_date(datum, now)

        key = str(bedrijf).strip().lower()
        if not key:
            skipped += 1
            continue

        contacts.setdefault(key, []).append((datum, str(type_raw)))

    wb.close()

    if skipped:
        print(f"[RECENT_CONTACTS] {skipped} rijen overgeslagen (leeg of ongeldig).")

    return contacts


def is_recent_contact(
    company: str,
    recent_contacts: dict[str, list[tuple[datetime, str]]],
    now: datetime | None = None,
) -> tuple[bool, str]:
    """
    Controleer of een bedrijf momenteel in cooldown zit.

    Returns:
        (True, reden)   als het bedrijf geblokkeerd is
        (False, "")     als het bedrijf benaderd mag worden
    """
    if now is None:
        now = datetime.now()

    key = company.strip().lower()

    # Fuzzy match: check ook of de key een substring is van een bedrijfsnaam in de dict
    # (voor kleine spellingsverschillen)
    matches = []
    if key in recent_contacts:
        matches = recent_contacts[key]
    else:
        # Probeer of de company-naam voorkomt als substring in een sleutel of vice versa
        for stored_key, rows in recent_contacts.items():
            if len(key) >= 4 and (key in stored_key or stored_key in key):
                matches.extend(rows)

    if not matches:
        return False, ""

    for datum, type_raw in matches:
        types   = _normalize_type(type_raw)
        age     = now - datum
        light   = _is_light_contact(types)
        cooldown = _COOLDOWN_LIGHT if light else _COOLDOWN_HEAVY

        if age < cooldown:
            contact_label = "gemaild/gebeld" if light else "zwaarder contact"
            days_remaining = (datum + cooldown - now).days
            datum_str = datum.strftime("%d-%m-%Y")
            return (
                True,
                f"{contact_label} op {datum_str} "
                f"({'3 mnd' if light else '1 jaar'} cooldown, "
                f"nog {days_remaining} dag(en))"
            )

    return False, ""


def get_blocked_companies(
    recent_contacts: dict[str, list[tuple[datetime, str]]],
    now: datetime | None = None,
) -> dict[str, str]:
    """
    Geef alle momenteel geblokkeerde bedrijven terug als {naam_lower: reden}.
    Handig voor batch-checks.
    """
    if now is None:
        now = datetime.now()

    blocked = {}
    for key, rows in recent_contacts.items():
        is_blocked, reason = is_recent_contact(key, recent_contacts, now)
        if is_blocked:
            blocked[key] = reason
    return blocked