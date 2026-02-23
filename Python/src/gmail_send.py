"""
gmail_send.py — E-mail verzending via Gmail API

Verantwoordelijkheden:
  - E-mailbericht aanmaken (MIME)
  - Versturen via Gmail API
  - Rate limiting helpers
"""

from __future__ import annotations

import base64
import time
from email.message import EmailMessage
from typing import Any


def create_message(
    sender_name: str,
    to_addr: str,
    subject: str,
    body_text: str,
) -> dict[str, str]:
    """
    Maak een Gmail API message dict aan vanuit plain-text parameters.

    Args:
        sender_name: Naam die in de Van-header verschijnt
        to_addr: Ontvanger e-mailadres
        subject: Onderwerpregel
        body_text: Plain-text body (newlines worden gerespecteerd)

    Returns:
        Dict met 'raw' key voor de Gmail API
    """
    msg = EmailMessage()
    msg["To"]      = to_addr
    msg["Subject"] = subject
    if sender_name:
        msg["From"] = sender_name   # Gmail gebruikt het ingelogde account, naam is display only

    msg.set_content(body_text)

    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode("utf-8")
    return {"raw": raw}


def send_one(service: Any, user_id: str, message: dict[str, str]) -> str:
    """
    Verstuur één bericht via de Gmail API.

    Args:
        service: Geauthenticeerde Gmail API service (van gmail_auth.get_gmail_service)
        user_id: Meestal "me" (het ingelogde account)
        message: Dict met 'raw' key (van create_message)

    Returns:
        Gmail message ID als string

    Raises:
        googleapiclient.errors.HttpError: Bij API fouten
    """
    resp = service.users().messages().send(userId=user_id, body=message).execute()
    return str(resp.get("id", ""))


def send_with_retry(
    service: Any,
    user_id: str,
    message: dict[str, str],
    max_retries: int = 3,
    retry_delay: float = 5.0,
) -> str:
    """
    Verstuur met automatische retry bij tijdelijke fouten (429, 503).

    Returns:
        Gmail message ID

    Raises:
        Exception: Na alle retries uitgeput
    """
    from googleapiclient.errors import HttpError

    last_error = None
    for attempt in range(1, max_retries + 1):
        try:
            return send_one(service, user_id, message)
        except HttpError as e:
            last_error = e
            status = e.resp.status if e.resp else 0

            if status in (429, 500, 503) and attempt < max_retries:
                wait = retry_delay * attempt
                print(f"[GMAIL] HTTP {status} — retry {attempt}/{max_retries} over {wait:.0f}s...")
                time.sleep(wait)
            else:
                raise

    raise last_error