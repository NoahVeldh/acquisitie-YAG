from __future__ import annotations

import base64
from email.message import EmailMessage
from typing import Any, Dict

import pandas as pd


def create_message(sender_name: str, to_addr: str, subject: str, body_text: str) -> Dict[str, str]:
    msg = EmailMessage()
    # Let op: Gmail API bepaalt de daadwerkelijke From (account), maar naam kan via headers.
    msg["To"] = to_addr
    msg["Subject"] = subject
    if sender_name:
        msg["From"] = sender_name

    msg.set_content(body_text)

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
    return {"raw": raw}


def render_email_body(row: pd.Series) -> str:
    # Heel basic template; pas aan aan jouw tone of voice
    first = (row.get("First Name") or "").strip()
    company = (row.get("Company") or "").strip()
    title = (row.get("Title") or "").strip()

    greet = f"Hoi {first}," if first else "Hoi,"
    line2 = f"Ik zag dat je werkzaam bent bij {company}." if company else "Ik kwam je profiel tegen."
    line3 = f"In jouw rol als {title} leek het me interessant om even kennis te maken." if title else "Het leek me interessant om even kennis te maken."

    return "\n".join(
        [
            greet,
            "",
            line2,
            line3,
            "",
            "Heb je deze week 10 minuten voor een korte kennismaking?",
            "",
            "Groet,",
            "Noah",
        ]
    )


def send_one(service: Any, user_id: str, message: Dict[str, str]) -> str:
    """Returns message id."""
    resp = service.users().messages().send(userId=user_id, body=message).execute()
    return str(resp.get("id", ""))


def subject_from_template(template: str, row: pd.Series) -> str:
    company = (row.get("Company") or "").strip()
    # veilige format
    return template.format(company=company or "jullie")
