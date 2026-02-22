from __future__ import annotations

import os
from typing import Sequence

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


DEFAULT_SCOPES: Sequence[str] = ("https://www.googleapis.com/auth/gmail.send",)


def get_gmail_service(
    credentials_json: str,
    token_json: str,
    scopes: Sequence[str] = DEFAULT_SCOPES,
):
    """
    Returns Gmail API service.
    - credentials_json: OAuth client secrets file (downloaded from Google Cloud Console)
    - token_json: saved token cache (will be created/updated)
    """
    creds = None
    if os.path.exists(token_json):
        creds = Credentials.from_authorized_user_file(token_json, scopes)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(credentials_json, scopes)
            creds = flow.run_local_server(port=0)

        os.makedirs(os.path.dirname(token_json), exist_ok=True)
        with open(token_json, "w", encoding="utf-8") as token:
            token.write(creds.to_json())

    service = build("gmail", "v1", credentials=creds)
    return service
