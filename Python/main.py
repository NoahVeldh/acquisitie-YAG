from __future__ import annotations

import os

import pandas as pd
from dotenv import load_dotenv

from src.acqlist import load_leads_from_excel, LeadColumns
from src.gmail_auth import get_gmail_service
from src.gmail_send import create_message, render_email_body, subject_from_template, send_one
from src.storage import load_suppression, append_suppression, append_send_log, load_do_not_contact, is_do_not_contact


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

    xlsx_path = os.getenv("ACQ_XLSX_PATH", "Python/data/Acquisitie.xlsx")
    suppression_path = os.getenv("SUPPRESSION_PATH", "Python/output/suppression.csv")
    send_log_path = os.getenv("SEND_LOG_PATH", "Python/output/send_log.csv")

    credentials_json = os.getenv("CREDENTIALS_JSON", "Python/credentials/credentials.json")
    token_json = os.getenv("TOKEN_JSON", "Python/credentials/token.json")

    dry_run = env_bool("DRY_RUN", True)
    max_emails = env_int("MAX_EMAILS", 20)

    sender_name = os.getenv("SENDER_NAME", "")
    subject_template = os.getenv("SUBJECT_TEMPLATE", "Even kennismaken — {company}")

    # 1) Load leads
    df = load_leads_from_excel(xlsx_path, sheet_name='Sheet1', columns=LeadColumns())
    if df.empty:
        print("Geen leads gevonden met geldige email_primary.")
        return

    # 2) suppression
    suppressed = load_suppression(suppression_path)
    dnc_path = os.getenv("DNC_PATH", "Python/data/Niet Benaderen.xlsx")
    do_not_contact = load_do_not_contact(dnc_path)   # <-- gooit error als bestand mist

    # 3) auth/service (alleen als niet dry-run)
    service = None
    if not dry_run:
        service = get_gmail_service(credentials_json=credentials_json, token_json=token_json)

    # 4) send loop
    sent_count = 0
    for _, row in df.iterrows():
        email = (row.get("email_primary") or "").strip().lower()
        company = str(row.get("Company", "")).strip()

        # ✅ STAP 1: niet-benaderen check — vóór alles
        blocked, reason = is_do_not_contact(company, do_not_contact)
        if blocked:
            print(f"[DNC] Geblokkeerd: {company!r} ({email}) — matched: {reason!r}")
            continue

        # Stap 2: basis email validatie
        if not email or "@" not in email:
            continue

        # Stap 3: email suppressie
        if email in suppressed:
            continue

        subject = subject_from_template(subject_template, row)
        body = render_email_body(row)

        if dry_run:
            print(f"[DRY_RUN] Would send to {email} | subject='{subject}'")
            append_send_log(
                send_log_path,
                {
                    "email": email,
                    "company": row.get("Company", ""),
                    "title": row.get("Title", ""),
                    "status": "DRY_RUN",
                    "message_id": "",
                    "error": "",
                    "subject": subject,
                    "body" : body,
                },
            )
        else:
            try:
                msg = create_message(sender_name=sender_name, to_addr=email, subject=subject, body_text=body)
                message_id = send_one(service, user_id="me", message=msg)
                append_suppression(suppression_path, email)
                append_send_log(
                    send_log_path,
                    {
                        "email": email,
                        "company": row.get("Company", ""),
                        "title": row.get("Title", ""),
                        "status": "SENT",
                        "message_id": message_id,
                        "error": "",
                    },
                )
            except Exception as e:
                append_send_log(
                    send_log_path,
                    {
                        "email": email,
                        "company": row.get("Company", ""),
                        "title": row.get("Title", ""),
                        "status": "ERROR",
                        "message_id": "",
                        "error": str(e),
                    },
                )

        sent_count += 1
        if sent_count >= max_emails:
            break

    print(f"Klaar. Verwerkt: {sent_count} (dry_run={dry_run}). Log: {send_log_path}")


if __name__ == "__main__":
    main()
