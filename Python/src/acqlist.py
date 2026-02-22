from __future__ import annotations

import re
from dataclasses import dataclass
from typing import Iterable, Optional

import pandas as pd


EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}")


@dataclass(frozen=True)
class LeadColumns:
    # Pas deze aan op jouw Excel-kolomnamen
    first_name: str = "First Name"
    last_name: str = "Last Name"
    company: str = "Company"
    title: str = "Title"
    email: str = "Email"  # als je in je sheet een eigen email-kolom hebt
    website: str = "Website"


def _extract_emails(value: object) -> list[str]:
    if value is None:
        return []
    text = str(value)
    return list(dict.fromkeys(EMAIL_RE.findall(text)))  # unique behoud volgorde


def load_leads_from_excel(
    xlsx_path: str,
    sheet_name: Optional[str] = None,
    columns: LeadColumns = LeadColumns(),
) -> pd.DataFrame:
    df = pd.read_excel(xlsx_path, sheet_name=sheet_name, engine="openpyxl")

    # Zorg dat basis kolommen bestaan; als niet, laat df gewoon door (dan kun je later mappen)
    # Maar we normaliseren alvast.
    for col in [columns.first_name, columns.last_name, columns.company, columns.title, columns.email, columns.website]:
        if col not in df.columns:
            df[col] = ""

    # Email normalisatie: soms staan er meerdere e-mails in één cel.
    df["_emails"] = df[columns.email].apply(_extract_emails)
    df["email_primary"] = df["_emails"].apply(lambda xs: xs[0] if xs else "")

    # Schone strings
    for col in [columns.first_name, columns.last_name, columns.company, columns.title, columns.website]:
        df[col] = df[col].fillna("").astype(str).str.strip()

    df["email_primary"] = df["email_primary"].fillna("").astype(str).str.strip().str.lower()

    # Filter: alleen rijen met een geldig email_primary
    df = df[df["email_primary"].str.contains("@", na=False)].copy()

    return df


def iter_recipients(df: pd.DataFrame) -> Iterable[str]:
    """Yield unieke emails (lowercased) uit email_primary."""
    seen = set()
    for email in df["email_primary"].astype(str):
        e = email.strip().lower()
        if not e or "@" not in e:
            continue
        if e in seen:
            continue
        seen.add(e)
        yield e
