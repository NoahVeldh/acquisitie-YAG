from __future__ import annotations

import os
import time

import pandas as pd
from dotenv import load_dotenv

from src.acqlist import load_leads_from_excel, LeadColumns
from src.gmail_auth import get_gmail_service
from src.gmail_send import create_message, render_email_body, subject_from_template, send_one
from src.storage import (
    load_suppression,
    append_suppression,
    append_send_log,
    load_do_not_contact,
    is_do_not_contact,
    load_contacted_companies,
)


def env_bool(name: str, default: bool = False) -> bool:
    v = os.getenv(name)
    if v is None:
        return default
    return v.strip().lower() in {"1", "true", "yes", "y", "on"}


def env_int(name: str, default: int) -> int:
    v = os.getenv(name)
    if v is None:
        return default
    try:
        return int(v)
    except ValueError:
        return default


def main() -> None:
    load_dotenv()

    # ── Paden ──────────────────────────────────────────────────────────────
    xlsx_path         = os.getenv("ACQ_XLSX_PATH",     "data/Acquisitie.xlsx")
    suppression_path  = os.getenv("SUPPRESSION_PATH",  "output/suppression.csv")
    send_log_path     = os.getenv("SEND_LOG_PATH",     "output/send_log.csv")
    credentials_json  = os.getenv("CREDENTIALS_JSON",  "credentials/credentials.json")
    token_json        = os.getenv("TOKEN_JSON",         "credentials/token.json")
    dnc_path          = os.getenv("DNC_PATH",           "data/Niet Benaderen - Alle Niet Benaderen.xlsx")

    # ── Consultant instellingen (pas aan in .env per persoon) ──────────────
    sender_name       = os.getenv("SENDER_NAME",        "")        # Naam die in de mail staat
    sender_email      = os.getenv("SENDER_EMAIL",       "")        # Gmail account waarmee verstuurd wordt
    subject_template  = os.getenv("SUBJECT_TEMPLATE",   "Young Advisory Group x {company}")

    # ── Run instellingen ───────────────────────────────────────────────────
    dry_run           = env_bool("DRY_RUN", True)
    max_emails        = env_int("MAX_EMAILS", 20)
    rate_limit_sec    = float(os.getenv("RATE_LIMIT_SEC", "2"))    # seconden tussen mails
    ai_body_column    = os.getenv("AI_BODY_COLUMN", "AI bericht")  # kolomnaam met AI-tekst

    # ── Kolomnamen Excel (pas aan als jouw sheet andere namen heeft) ───────
    columns = LeadColumns(
        first_name = os.getenv("COL_FIRST_NAME", "First Name"),
        last_name  = os.getenv("COL_LAST_NAME",  "Last Name"),
        company    = os.getenv("COL_COMPANY",    "Company Name"),
        title      = os.getenv("COL_TITLE",      "Job Title"),
        email      = os.getenv("COL_EMAIL",      "Email(s)"),
        website    = os.getenv("COL_WEBSITE",    "Company Website"),
    )

    # 1) Leads laden
    df = load_leads_from_excel(xlsx_path, sheet_name='Sheet1', columns=columns)
    if df.empty:
        print("Geen leads gevonden met geldige email_primary.")
        return

    # 2) Suppressie & DNC laden
    suppressed          = load_suppression(suppression_path)
    contacted_companies = load_contacted_companies(send_log_path)
    do_not_contact      = load_do_not_contact(dnc_path)

    # 3) Gmail service (alleen bij echte send)
    service = None
    if not dry_run:
        service = get_gmail_service(credentials_json=credentials_json, token_json=token_json)

    # 4) Send loop
    sent_count   = 0
    skip_dnc     = 0
    skip_sup     = 0
    skip_company = 0
    skip_no_body = 0

    for _, row in df.iterrows():
        email   = (row.get("email_primary") or "").strip().lower()
        company = str(row.get(columns.company, "")).strip()

        # Check 1: DNC lijst
        blocked, reason = is_do_not_contact(company, do_not_contact)
        if blocked:
            print(f"[DNC CHECK]      Geblokkeerd: {company!r} ({email}) — matched: {reason!r}")
            skip_dnc += 1
            continue

        # Check 2: email validatie
        if not email or "@" not in email:
            continue

        # Check 3: email al eerder gestuurd
        if email in suppressed:
            skip_sup += 1
            continue

        # Check 4: bedrijf al eerder benaderd (andere collega/contact)
        if company.lower().strip() in contacted_companies:
            print(f"[SKIP]     Bedrijf al benaderd: {company!r} ({email})")
            skip_company += 1
            continue

        # Body ophalen uit AI bericht kolom
        body = str(row.get(ai_body_column) or "").strip()
        if not body:
            print(f"[SKIP]     Geen AI bericht voor {email}, sla over.")
            skip_no_body += 1
            continue

        subject = subject_template.format(company=company or "jullie")

        if dry_run:
            print(f"[DRY_RUN]  To: {email}")
            print(f"           Subject: {subject}")
            append_send_log(
                    send_log_path,
                    {
                        "email":      email,
                        "company":    company,
                        "title":      row.get(columns.title, ""),
                        "consultant": sender_name,
                        "status":     "DRY RUN",
                        "message_id": '',
                        "error":      "",
                        "subject":    subject,
                        "body":       body,
                    },
            )
        else:
            try:
                msg = create_message(
                    sender_name=sender_name,
                    to_addr=email,
                    subject=subject,
                    body_text=body,
                )
                message_id = send_one(service, user_id="me", message=msg)
                append_suppression(suppression_path, email)
                append_send_log(
                    send_log_path,
                    {
                        "email":      email,
                        "company":    company,
                        "title":      row.get(columns.title, ""),
                        "consultant": sender_name,
                        "status":     "SENT",
                        "message_id": message_id,
                        "error":      "",
                        "subject":    subject,
                        "body":       body,
                    },
                )
                print(f"[SENT]     {email} | {company}")

                # Rate limiting — voorkom spam-detectie
                time.sleep(rate_limit_sec)

            except Exception as e:
                print(f"[ERROR]    {email} — {e}")
                append_send_log(
                    send_log_path,
                    {
                        "email":      email,
                        "company":    company,
                        "title":      row.get(columns.title, ""),
                        "consultant": sender_name,
                        "status":     "ERROR",
                        "message_id": "",
                        "error":      str(e),
                        "subject":    subject,
                        "body":       body,
                    },
                )

        sent_count += 1
        if sent_count >= max_emails:
            break

    # Samenvatting
    print(f"""
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Klaar! (dry_run={dry_run})
  Verwerkt:          {sent_count}
  Geblokkeerd (DNC): {skip_dnc}
  Al gemaild:        {skip_sup}
  Bedrijf al benaderd: {skip_company}
  Geen AI bericht:   {skip_no_body}
  Log: {send_log_path}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━""")


if __name__ == "__main__":
    main()
