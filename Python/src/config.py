"""
config.py — Centrale configuratie voor YAG Acquisitie Tool

Alle kolom-indices staan hier gedefinieerd als één enkele source of truth.
Pas alleen hier aan als de sheet-structuur verandert.

Sheet kolom layout:
  A=1   Company
  B=2   First Name
  C=3   Last Name
  D=4   Job Title
  E=5   Email
  F=6   Phone
  G=7   LinkedIn URL
  H=8   Enriched ✅
  I=9   AI Status
  J=10  Mail Status
  K=11  Datum Mail
  L=12  Follow-up datum
  M=13  Reactie ontvangen
  N=14  Opmerking
  O=15  --- separator ---
  P=16  Consultant
  Q=17  Vestiging
  R=18  Type
  S=19  Gevallen
  T=20  Hoe kom je aan dit contact
  U=21  --- separator ---
  V=22  Request ID
  W=23  Contact ID
  X=24  isShown
  Y=25  AI Bericht
"""

from __future__ import annotations

# ── Sheet kolom nummers (1-indexed, zoals gspread verwacht) ───────────────
class Col:
    COMPANY           = 1
    FIRST_NAME        = 2
    LAST_NAME         = 3
    JOB_TITLE         = 4
    EMAIL             = 5
    PHONE             = 6
    LINKEDIN_URL      = 7
    ENRICHED          = 8
    AI_STATUS         = 9
    MAIL_STATUS       = 10
    DATUM_MAIL        = 11
    FOLLOWUP_DATUM    = 12
    REACTIE           = 13
    OPMERKING         = 14
    # O=15 separator
    CONSULTANT        = 16
    VESTIGING         = 17
    TYPE              = 18
    GEVALLEN          = 19
    HOE_CONTACT       = 20
    # U=21 separator
    REQUEST_ID        = 22
    CONTACT_ID        = 23
    IS_SHOWN          = 24
    AI_BERICHT        = 25

    TOTAL_COLS        = 25

# ── Kolomletter helpers (voor foutmeldingen / logging) ────────────────────
def col_letter(n: int) -> str:
    """Zet 1-indexed kolomnummer om naar letter (1→A, 26→Z, 27→AA)."""
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


# ── Status waarden ─────────────────────────────────────────────────────────
class AIStatus:
    PENDING  = "PENDING"
    RUNNING  = "RUNNING"
    DONE     = "✅ DONE"
    ERROR    = "❌ ERROR"
    SKIPPED  = "⏭ SKIPPED"          # geen website / te weinig info


class MailStatus:
    PENDING  = "PENDING"
    DRY_RUN  = "DRY RUN"
    SENT     = "✅ SENT"
    ERROR    = "❌ ERROR"
    DNC      = "🚫 DNC"              # Do Not Contact
    SUPPRESSED = "⏭ AL GEMAILD"
    NO_EMAIL = "⚠ GEEN EMAIL"


class Enriched:
    YES = "✅ Yes"
    NO  = "No"


# ── Verplichte consultant-velden (vóór AI generatie te vullen) ────────────
REQUIRED_META_COLS = [Col.CONSULTANT, Col.VESTIGING, Col.TYPE, Col.HOE_CONTACT]
REQUIRED_META_NAMES = ["Consultant", "Vestiging", "Type", "Hoe contact"]

# ── Data rij start (rij 1 = header) ──────────────────────────────────────
DATA_START_ROW = 2