"""
gmail_auth.py — Gmail OAuth2 authenticatie

Gebruikt OAuth2 met een lokaal opgeslagen token (token.json).
De eerste keer opent dit een browser voor toestemming.
Daarna wordt het token automatisch vernieuwd.
"""

from __future__ import annotations

import os
from typing import Sequence

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

DEFAULT_SCOPES: Sequence[str] = ("https://www.googleapis.com/auth/gmail.send",)


def get_gmail_service(
    credentials_json: str,
    token_json: str,
    scopes: Sequence[str] = DEFAULT_SCOPES,
):
    """
    Geeft een geauthenticeerde Gmail API service terug.

    Args:
        credentials_json: OAuth client secrets (download via Google Cloud Console)
        token_json: Pad waar het token opgeslagen wordt/is
        scopes: Benodigde API scopes

    Eerste keer gebruik:
      1. Download credentials.json via Google Cloud Console
         → APIs & Services → Credentials → OAuth 2.0 Client ID → Download
      2. Zet het bestand op credentials/credentials.json
      3. Run het script → browser opent voor toestemming
      4. token.json wordt automatisch opgeslagen voor volgende runs
    """
    if not os.path.exists(credentials_json):
        raise FileNotFoundError(
            f"credentials.json niet gevonden: {credentials_json}\n"
            "Download via Google Cloud Console → APIs & Services → Credentials."
        )

    creds = None

    if os.path.exists(token_json):
        creds = Credentials.from_authorized_user_file(token_json, scopes)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            print("[GMAIL] Token verlopen, vernieuwen...")
            creds.refresh(Request())
        else:
            print("[GMAIL] Eerste authenticatie — browser opent voor toestemming...")
            flow = InstalledAppFlow.from_client_secrets_file(credentials_json, scopes)
            creds = flow.run_local_server(port=0)

        os.makedirs(os.path.dirname(token_json), exist_ok=True)
        with open(token_json, "w", encoding="utf-8") as token:
            token.write(creds.to_json())
        print(f"[GMAIL] Token opgeslagen: {token_json}")

    service = build("gmail", "v1", credentials=creds)
    print("[GMAIL] ✅ Geauthenticeerd.")
    return service