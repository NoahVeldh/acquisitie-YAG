"""
gmail_send.py — E-mail verzending via SMTP met Gmail App Password

Verstuurt HTML-mails met:
  - Logo uit een lokaal bestand (PNG aanbevolen — zie LOGO_PATH in .env)
  - E-mailadres in handtekening als klikbare mailto: link
  - LinkedIn URL als klikbare link (stel in via SENDER_LINKEDIN in .env)
  - Volledige breedte — past zich aan aan het scherm van de ontvanger
  - Plain-text fallback voor clients die geen HTML ondersteunen
  - Optionele PDF bijlage (stel in via ATTACHMENT_PDF in .env)

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
LOGO INSTELLEN

  Aanbevolen formaat: PNG
    - Lossless kwaliteit, transparantie-ondersteuning
    - Werkt in Gmail, Outlook, Apple Mail en alle grote clients
    - Houd het bestand onder ~100 kB voor snelle weergave

  Stap 1: Zet je logo-bestand in de map assets/
          Voorbeeld: assets/logo.png

  Stap 2: Voeg toe aan consultants/<naam>.env:
          LOGO_PATH=assets/logo.png

  Geen LOGO_PATH ingesteld → handtekening verschijnt zonder logo.
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

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
from email.mime.application import MIMEApplication
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 587


def _load_logo(logo_path: str) -> tuple[bytes, str] | tuple[None, None]:
    """
    Laad logo uit bestand. Ondersteunt PNG en JPG/JPEG.
    Geeft (bytes, subtype) terug, of (None, None) als het bestand niet bestaat.
    """
    if not logo_path:
        return None, None

    path = Path(logo_path)
    if not path.exists():
        print(f"[GMAIL] ⚠ Logo niet gevonden: {logo_path} — mail verstuurd zonder logo")
        return None, None

    suffix = path.suffix.lower()
    subtype_map = {".png": "png", ".jpg": "jpeg", ".jpeg": "jpeg"}
    subtype = subtype_map.get(suffix)
    if not subtype:
        print(f"[GMAIL] ⚠ Onbekend logo formaat '{suffix}' — gebruik .png of .jpg. Mail verstuurd zonder logo.")
        return None, None

    return path.read_bytes(), subtype


def _text_to_html(text: str) -> str:
    """
    Zet plain-text e-mailbody om naar HTML.
    Bewaart alinea-scheidingen (dubbele enters) en enkelvoudige regeleinden.
    """
    import html as html_lib
    paragraphs = text.split("\n\n")
    parts = []
    for para in paragraphs:
        escaped = html_lib.escape(para)
        escaped = escaped.replace("\n", "<br>")
        parts.append(f"<p>{escaped}</p>")
    return "\n".join(parts)


def build_html_email(
    body_text: str,
    sender_name: str,
    sender_email: str,
    sender_phone: str,
    sender_linkedin: str,
    vestiging: str,
    has_logo: bool = False,
) -> str:
    """
    Bouw de volledige HTML e-mail op.
    De body_text bevat de mail inclusief plain-text handtekening.
    Wij knippen de handtekening eraf en vervangen hem door HTML.
    """
    split_marker = "Met vriendelijke groet,"
    if split_marker in body_text:
        mail_body = body_text[:body_text.index(split_marker)].strip()
    else:
        mail_body = body_text.strip()

    body_html = _text_to_html(mail_body)

    # LinkedIn URL — zorg dat er altijd https:// voor staat
    linkedin_url = sender_linkedin.strip()
    if linkedin_url and not linkedin_url.startswith("http"):
        linkedin_url = "https://" + linkedin_url

    # Logo regel — alleen renderen als het logo geladen is
    logo_html = (
        """      <tr>
        <td style="padding-bottom:8px;">
          <img src="cid:yag_logo" alt="YAG" height="48" style="display:block;">
        </td>
      </tr>"""
        if has_logo else ""
    )

    html = f"""<!DOCTYPE html>
<html lang="nl">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<body style="margin:0;padding:0;font-family:Arial,sans-serif;font-size:14px;color:#222222;line-height:1.6;">

  <div style="width:100%;max-width:100%;">

    <!-- Mail body -->
    <div style="margin-bottom:24px;">
      {body_html}
    </div>

    <!-- Handtekening -->
    <table cellpadding="0" cellspacing="0" border="0" style="font-family:Arial,sans-serif;font-size:13px;color:#222222;">
      <tr>
        <td style="padding-bottom:2px;">
          <strong>{sender_name}</strong>
        </td>
      </tr>
      <tr>
        <td style="padding-bottom:8px;color:#555555;">
          Strategy Consultant - Young Advisory Group
        </td>
      </tr>
{logo_html}
      <tr>
        <td style="padding-bottom:6px;color:#999999;font-size:12px;">
          &ndash;&ndash;&ndash;&ndash;&ndash;&ndash;&ndash;&ndash;&ndash;&ndash;&ndash;&ndash;&ndash;&ndash;&ndash;&ndash;&ndash;&ndash;&ndash;&ndash;&ndash;&ndash;&ndash;&ndash;&ndash;
        </td>
      </tr>
      <tr><td style="color:#555555;">Vestiging {vestiging}</td></tr>
      <tr><td style="color:#555555;">Videolab</td></tr>
      <tr><td style="color:#555555;">Torenallee 20</td></tr>
      <tr><td style="color:#555555;">5617 BC Eindhoven</td></tr>
      <tr><td style="color:#555555;padding-top:4px;">{sender_phone}</td></tr>
      <tr>
        <td style="padding-top:4px;color:#555555;">
          <a href="mailto:{sender_email}" style="color:#555555;text-decoration:none;">{sender_email}</a>
          &nbsp;|&nbsp;
          <a href="{linkedin_url}" style="color:#3dbbac;text-decoration:none;">LinkedIn</a>
          &nbsp;|&nbsp;
          <a href="https://www.youngadvisorygroup.nl" style="color:#3dbbac;text-decoration:none;">www.youngadvisorygroup.nl</a>
        </td>
      </tr>
    </table>

  </div>
</body>
</html>"""
    return html


def send_email(
    sender_email: str,
    sender_name: str,
    app_password: str,
    to_addr: str,
    subject: str,
    body_text: str,
    sender_phone: str = "",
    sender_linkedin: str = "",
    vestiging: str = "",
    attachment_pdf: str = "",
    logo_path: str = "",
    max_retries: int = 3,
    retry_delay: float = 5.0,
) -> str:
    """
    Verstuur één HTML e-mail via Gmail SMTP met App Password.

    Args:
        sender_email:    Gmail adres waarmee verstuurd wordt
        sender_name:     Naam die in Van-header verschijnt
        app_password:    16-tekens App Password van Google
        to_addr:         Ontvanger e-mailadres
        subject:         Onderwerpregel
        body_text:       Plain-text body (ook als fallback)
        sender_phone:    Telefoonnummer voor handtekening
        sender_linkedin: LinkedIn profiel URL (stel in via SENDER_LINKEDIN)
        vestiging:       Vestiging naam voor handtekening
        attachment_pdf:  Pad naar PDF bijlage (leeg = geen bijlage)
        logo_path:       Pad naar logo bestand PNG/JPG (leeg = geen logo)
        max_retries:     Aantal pogingen bij tijdelijke fouten
        retry_delay:     Seconden wachten tussen retries

    Returns:
        Een unieke message ID string
    """
    if not sender_email:
        raise ValueError("sender_email is verplicht")
    if not app_password:
        raise ValueError(
            "GMAIL_APP_PASSWORD ontbreekt in .env\n"
            "Maak een App Password aan via myaccount.google.com/apppasswords"
        )

    # ── Logo laden ────────────────────────────────────────────────────────
    logo_bytes, logo_subtype = _load_logo(logo_path)
    has_logo = logo_bytes is not None

    # ── HTML bouwen ───────────────────────────────────────────────────────
    html_body = build_html_email(
        body_text=body_text,
        sender_name=sender_name,
        sender_email=sender_email,
        sender_phone=sender_phone,
        sender_linkedin=sender_linkedin,
        vestiging=vestiging,
        has_logo=has_logo,
    )

    # multipart/related bevat HTML + ingesloten afbeeldingen (logo)
    msg_related = MIMEMultipart("related")

    msg_alternative = MIMEMultipart("alternative")
    msg_alternative.attach(MIMEText(body_text, "plain", "utf-8"))
    msg_alternative.attach(MIMEText(html_body, "html", "utf-8"))
    msg_related.attach(msg_alternative)

    # Logo als CID embed (alleen als het bestand bestaat)
    if has_logo:
        logo_filename = Path(logo_path).name
        logo_img = MIMEImage(logo_bytes, _subtype=logo_subtype)
        logo_img.add_header("Content-ID", "<yag_logo>")
        logo_img.add_header("Content-Disposition", "inline", filename=logo_filename)
        msg_related.attach(logo_img)

    # Buitenste envelop — mixed zodat bijlagen kunnen worden toegevoegd
    msg = MIMEMultipart("mixed")
    msg["Subject"] = subject
    msg["From"]    = f"{sender_name} <{sender_email}>" if sender_name else sender_email
    msg["To"]      = to_addr
    msg.attach(msg_related)

    # ── PDF bijlage ───────────────────────────────────────────────────────
    if attachment_pdf:
        pdf_path = Path(attachment_pdf)
        if pdf_path.exists():
            with open(pdf_path, "rb") as f:
                pdf_bytes = f.read()
            pdf_part = MIMEApplication(pdf_bytes, _subtype="pdf")
            pdf_part.add_header(
                "Content-Disposition",
                "attachment",
                filename=pdf_path.name,
            )
            msg.attach(pdf_part)
            print(f"[GMAIL] 📎 Bijlage toegevoegd: {pdf_path.name} ({len(pdf_bytes) / 1024:.0f} kB)")
        else:
            print(f"[GMAIL] ⚠ PDF bijlage niet gevonden, mail verstuurd zonder bijlage: {attachment_pdf}")

    # ── SMTP verzending met retries ───────────────────────────────────────
    import time as _time
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
            raw = f"{sender_email}{to_addr}{subject}{_time.time()}"
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
                _time.sleep(retry_delay * attempt)
            else:
                raise

    raise last_error


def verify_connection(sender_email: str, app_password: str) -> bool:
    """
    Test of de SMTP verbinding en het App Password werken.
    Returns True als verbinding en login lukken, anders False.
    """
    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=10) as server:
            server.ehlo()
            server.starttls()
            server.login(sender_email, app_password.replace(" ", ""))
        return True
    except Exception:
        return False