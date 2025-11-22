from typing import List, Dict, Any
from pathlib import Path

from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from ibm_watsonx_orchestrate.agent_builder.tools import tool

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

# BASE_DIR = folder "gmail_tool"
BASE_DIR = Path(__file__).resolve().parents[1]
TOKEN_PATH = BASE_DIR / "token.json"


def _get_gmail_service():
    """Load token.json dari paket tool dan buat service Gmail."""
    creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
    service = build("gmail", "v1", credentials=creds)
    return service


@tool()
def list_recent_emails(query: str = "", max_results: int = 10) -> List[Dict[str, Any]]:
    """
    Mengambil daftar email terbaru dari Gmail.

    Args:
        query: Query Gmail (mis. 'subject:reimbursement OR subject:claim').
        max_results: Jumlah maksimum email yang diambil.

    Returns:
        List dict berisi id, subject, from, date, snippet.
    """
    service = _get_gmail_service()

    resp = (
        service.users()
        .messages()
        .list(userId="me", q=query, maxResults=max_results)
        .execute()
    )
    messages = resp.get("messages", [])

    results: List[Dict[str, Any]] = []

    for m in messages:
        msg = (
            service.users()
            .messages()
            .get(
                userId="me",
                id=m["id"],
                format="metadata",
                metadataHeaders=["Subject", "From", "Date"],
            )
            .execute()
        )

        headers = {h["name"]: h["value"] for h in msg["payload"]["headers"]}

        results.append(
            {
                "id": m["id"],
                "subject": headers.get("Subject", ""),
                "from": headers.get("From", ""),
                "date": headers.get("Date", ""),
                "snippet": msg.get("snippet", ""),
            }
        )

    return results
