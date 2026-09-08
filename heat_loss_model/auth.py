"""
Google Sheets authentication helper supporting Service Accounts and OAuth credentials.
"""

import os
import json
from pathlib import Path
from typing import Optional
import gspread
from google.oauth2.credentials import Credentials
from google.oauth2 import service_account
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def get_gspread_client(
    credentials_path: Optional[str] = None,
    service_account_path: Optional[str] = None,
    token_path: Optional[str] = None
) -> gspread.Client:
    """
    Authenticates and returns an authorized gspread Client.
    Priority:
    1. Service Account (service_account.json / env var)
    2. Local OAuth token.json
    3. OAuth Client ID (credentials.json) browser login
    """
    base_dir = Path(__file__).resolve().parent.parent

    # 1. Check for Service Account in env or file
    sa_file = service_account_path or os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE")
    if sa_file:
        resolved_sa = Path(sa_file)
        if not resolved_sa.is_absolute():
            resolved_sa = base_dir / sa_file
        if resolved_sa.exists():
            creds = service_account.Credentials.from_service_account_file(
                str(resolved_sa), scopes=SCOPES
            )
            return gspread.authorize(creds)
    
    default_sa_file = base_dir / "service_account.json"
    if default_sa_file.exists():
        creds = service_account.Credentials.from_service_account_file(
            str(default_sa_file), scopes=SCOPES
        )
        return gspread.authorize(creds)

    # 2. Check for local OAuth token.json
    tok_path = Path(token_path or base_dir / "token.json")
    creds = None
    if tok_path.exists():
        try:
            creds = Credentials.from_authorized_user_file(str(tok_path), SCOPES)
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
        except Exception:
            creds = None

    # 3. Standard credentials.json browser authorization
    if not creds or not creds.valid:
        cred_file = Path(credentials_path or base_dir / "credentials.json")
        if not cred_file.exists():
            raise FileNotFoundError(
                f"Google credentials not found.\n"
                f"Please place your 'service_account.json' or 'credentials.json' "
                f"in:\n  {base_dir}\n"
                f"Or set GOOGLE_SERVICE_ACCOUNT_FILE in .env."
            )
        flow = InstalledAppFlow.from_client_secrets_file(str(cred_file), SCOPES)
        creds = flow.run_local_server(port=0)
        with open(tok_path, "w", encoding="utf-8") as token:
            token.write(creds.to_json())

    return gspread.authorize(creds)
