"""
gmail_send.py — E-mail verzending via SMTP met Gmail App Password

Veel simpeler dan OAuth: geen credentials.json, geen token.json,
geen browser-login. Alleen een App Password nodig.

Hoe maak je een App Password aan:
  1. Ga naar myaccount.google.com/security
  2. Zet 2-stapsverificatie aan (verplicht)
  3. Ga naar myaccount.google.com/apppasswords
  4. App name: "yag-mailer" → Create
  5. Kopieer de 16 tekens → zet in consultants/<naam>.env als GMAIL_APP_PASSWORD
"""

from __future__ import annotations

import hashlib
import smtplib
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587


def send_email(
    sender_email: str,
    sender_name: str,
    app_password: str,
    to_addr: str,
    subject: str,
    body_text: str,
    max_retries: int = 3,
    retry_delay: float = 5.0,
) -> str:
    """
    Verstuur één e-mail via Gmail SMTP met App Password.

    Args:
        sender_email:  Gmail adres waarmee verstuurd wordt
        sender_name:   Naam die in Van-header verschijnt
        app_password:  16-tekens App Password van Google
        to_addr:       Ontvanger e-mailadres
        subject:       Onderwerpregel
        body_text:     Plain-text body
        max_retries:   Aantal pogingen bij tijdelijke fouten
        retry_delay:   Seconden wachten tussen retries

    Returns:
        Een unieke message ID string

    Raises:
        smtplib.SMTPAuthenticationError: Bij verkeerd wachtwoord
        ValueError: Als verplichte parameters ontbreken
    """
    if not sender_email:
        raise ValueError("sender_email is verplicht")
    if not app_password:
        raise ValueError(
            "GMAIL_APP_PASSWORD ontbreekt in .env\n"
            "Maak een App Password aan via myaccount.google.com/apppasswords"
        )

    # Bouw het MIME bericht
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"]    = f"{sender_name} <{sender_email}>" if sender_name else sender_email
    msg["To"]      = to_addr
    msg.attach(MIMEText(body_text, "plain", "utf-8"))

    # Verwijder spaties uit app password (Google toont het soms met spaties)
    password = app_password.replace(" ", "")

    last_error = None
    for attempt in range(1, max_retries + 1):
        try:
            with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30) as server:
                server.ehlo()
                server.starttls()
                server.ehlo()
                server.login(sender_email, password)
                server.sendmail(sender_email, to_addr, msg.as_string())

            # Genereer een simpele message ID
            raw = f"{sender_email}{to_addr}{subject}{time.time()}"
            return hashlib.md5(raw.encode()).hexdigest()[:16]

        except smtplib.SMTPAuthenticationError:
            raise smtplib.SMTPAuthenticationError(
                535,
                b"Gmail authenticatie mislukt. Controleer SENDER_EMAIL en GMAIL_APP_PASSWORD.\n"
                b"Zorg dat 2FA aan staat en gebruik een App Password, niet je gewone wachtwoord."
            )
        except (smtplib.SMTPServerDisconnected, smtplib.SMTPConnectError, OSError) as e:
            last_error = e
            if attempt < max_retries:
                print(f"[GMAIL] Verbindingsfout — retry {attempt}/{max_retries} over {retry_delay:.0f}s...")
                time.sleep(retry_delay * attempt)
            else:
                raise

    raise last_error


def verify_connection(sender_email: str, app_password: str) -> bool:
    """
    Test of de SMTP verbinding en het App Password werken.
    Handig om te checken vóórdat je een batch verstuurt.

    Returns:
        True als verbinding en login lukken, anders False
    """
    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=10) as server:
            server.ehlo()
            server.starttls()
            server.login(sender_email, app_password.replace(" ", ""))
        return True
    except Exception:
        return False