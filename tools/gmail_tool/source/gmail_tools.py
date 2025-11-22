from typing import List, Dict, Any
from pathlib import Path

from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from ibm_watsonx_orchestrate.agent_builder.tools import tool
from datetime import date, timedelta


from io import BytesIO
from base64 import urlsafe_b64decode
from PyPDF2 import PdfReader
import re
from pathlib import Path


SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

# BASE_DIR = folder "gmail_tool"
BASE_DIR = Path(__file__).resolve().parents[1]
TOKEN_PATH = BASE_DIR / "token.json"


def _get_gmail_service():
    """Load token.json dari paket tool dan buat service Gmail."""
    creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
    service = build("gmail", "v1", credentials=creds)
    return service

def _get_current_week_range_until_today():
    """
    Ambil range Senin hingga hari ini.
    - monday: Senin minggu ini
    - today_exclusive: besok (dipakai sebagai before: YYYY/MM/DD)
    """
    today = date.today()
    weekday = today.weekday()  # Monday=0

    monday = today - timedelta(days=weekday)
    tomorrow = today + timedelta(days=1)

    def fmt(d: date) -> str:
        return d.strftime("%Y/%m/%d")

    return fmt(monday), fmt(tomorrow)


def _extract_text_body(payload) -> str:
    """Ambil isi email (text/plain) dari struktur multipart Gmail."""
    if not payload:
        return ""
    mime = payload.get("mimeType", "")
    body = payload.get("body", {})

    if mime == "text/plain" and "data" in body:
        data = body["data"]
        return urlsafe_b64decode(data + "==").decode("utf-8", errors="ignore")

    for part in payload.get("parts", []):
        txt = _extract_text_body(part)
        if txt:
            return txt
    return ""

def _download_attachment_bytes(service, message_id: str, attachment_id: str) -> bytes:
    """Download attachment Gmail sebagai bytes."""
    att = (
        service.users()
        .messages()
        .attachments()
        .get(userId="me", messageId=message_id, id=attachment_id)
        .execute()
    )
    data = att.get("data", "")
    return urlsafe_b64decode(data.encode("utf-8"))

def _parse_reimburse_form_pdf(pdf_bytes: bytes) -> Dict[str, Any]:
    """
    Parse form reimburse Pituku (seperti contoh PDF) menjadi data terstruktur.
    Asumsi layout mirip dengan contoh: Tanggal Pengajuan, DETAIL PENGAJUAN, Nama Bank, dst.
    """
    reader = PdfReader(BytesIO(pdf_bytes))
    text_parts = []
    for page in reader.pages:
        t = page.extract_text() or ""
        text_parts.append(t)
    full_text = "\n".join(text_parts)

    # Normalisasi jadi list baris
    lines = [ln.strip() for ln in full_text.splitlines() if ln.strip()]

    # Tanggal Pengajuan
    tanggal = ""
    for ln in lines:
        if ln.startswith("Tanggal Pengajuan"):
            # contoh: "Tanggal Pengajuan : 16/04/2025"
            parts = ln.split(":", 1)
            if len(parts) == 2:
                tanggal = parts[1].strip()
            break

    # Nama Bank, No Rek, Nama Rek
    nama_bank = nomor_rek = nama_rek = ""
    for ln in lines:
        if ln.startswith("Nama Bank"):
            nama_bank = ln.split(":", 1)[1].strip()
        elif ln.startswith("Nomor Rekening"):
            nomor_rek = ln.split(":", 1)[1].strip()
        elif ln.startswith("Nama Rekening"):
            nama_rek = ln.split(":", 1)[1].strip()

    # Baris detail pengajuan
    # Kita cari baris yang mengandung "Fee Driver" dll, kemudian ambil angka-angkanya
    item_description = ""
    harga = jumlah = subtotal = None

    # gabungkan 3 baris setelah "No Description Detail Harga Jumlah Sub Total"
    for i, ln in enumerate(lines):
        if ln.startswith("No Description"):
            row_text = " ".join(lines[i + 1 : i + 4])  # relatif aman utk contoh
            # Deskripsi: buang "1 " di depan dan apapun setelah "Rp"
            m_desc = re.search(r"1\s+(.*?)\s+Rp", row_text)
            if m_desc:
                item_description = m_desc.group(1).strip()

            # Angka: ambil semua pola 300.000, 1.250.000, dst
            nums = re.findall(r"(\d{1,3}(?:\.\d{3})+)", row_text)
            def to_int(s: str) -> int:
                return int(s.replace(".", ""))

            if nums:
                harga = to_int(nums[0])
                subtotal = to_int(nums[-1])
            # Cari jumlah (biasanya 1, 2, dst) di antara harga & subtotal
            m_jumlah = re.search(r"\s(\d+)\s+Rp", row_text)
            if m_jumlah:
                jumlah = int(m_jumlah.group(1))

            break

    return {
        "tanggal_pengajuan": tanggal,
        "items": [
            {
                "description": item_description,
                "harga": harga,
                "jumlah": jumlah,
                "subtotal": subtotal,
            }
        ],
        "total": subtotal,
        "rekening": {
            "bank": nama_bank,
            "nomor_rekening": nomor_rek,
            "nama_rekening": nama_rek,
        },
    }


def _collect_attachments(payload, out_list):
    """Kumpulkan attachment (id, filename, mimeType) dari payload Gmail."""
    if not payload:
        return

    filename = payload.get("filename")
    body = payload.get("body", {})
    mime = payload.get("mimeType")

    if filename and body.get("attachmentId"):
        out_list.append(
            {
                "attachment_id": body["attachmentId"],
                "filename": filename,
                "mimeType": mime,
            }
        )

    for part in payload.get("parts", []):
        _collect_attachments(part, out_list)

@tool()
def list_reimburse_emails_this_week(max_results: int = 50) -> Dict[str, Any]:
    """
    Mengambil email reimbursement untuk periode Senin–hari ini.
    Dipakai untuk testing atau proses mingguan.
    """
    start_date, tomorrow = _get_current_week_range_until_today()
    return list_reimburse_emails_for_period(start_date, tomorrow, max_results)

@tool()
def list_reimburse_emails_for_period(start_date: str, end_date: str, max_results: int = 50) -> Dict[str, Any]:
    """
    Mengambil email reimbursement dalam rentang tanggal tertentu.

    Args:
        start_date: Tanggal awal (inclusive) format YYYY/MM/DD, contoh: '2025/04/14'.
        end_date: Tanggal akhir eksklusif (dipakai di 'before'), contoh: '2025/04/18'.
        max_results: Maks jumlah email yang diambil.

    Returns:
        Dict yang berisi info periode dan list email (id, subject, from, date).
    """
    service = _get_gmail_service()

    query = f'("reimburse" OR "reimbursement") after:{start_date} before:{end_date}'

    resp = (
        service.users()
        .messages()
        .list(userId="me", q=query, maxResults=max_results)
        .execute()
    )

    messages = resp.get("messages", [])

    emails: List[Dict[str, Any]] = []

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
        emails.append(
            {
                "id": m["id"],
                "subject": headers.get("Subject", ""),
                "from": headers.get("From", ""),
                "date": headers.get("Date", ""),
                "internal_ts": int(msg.get("internalDate", 0)),
            }
        )

    # sort newest first berdasarkan internalDate
    emails_sorted = sorted(emails, key=lambda x: x["internal_ts"], reverse=True)
    for e in emails_sorted:
        e.pop("internal_ts", None)

    return {
        "start_date": start_date,
        "end_date_exclusive": end_date,
        "count": len(emails_sorted),
        "emails": emails_sorted,
    }

@tool()
def parse_reimburse_form_from_email(message_id: str) -> Dict[str, Any]:
    """
    Cari attachment PDF pertama di email ini, download, lalu parse form reimburse.

    Args:
        message_id: ID email di Gmail.

    Returns:
        Dict berisi ringkasan reimburse (tanggal, item, total, info rekening).
    """
    service = _get_gmail_service()

    msg = (
        service.users()
        .messages()
        .get(userId="me", id=message_id, format="full")
        .execute()
    )

    attachments = []
    _collect_attachments(msg.get("payload"), attachments)

    # pilih attachment PDF pertama
    pdf_att = None
    for att in attachments:
        mime = (att.get("mimeType") or "").lower()
        filename = (att.get("filename") or "").lower()
        if "pdf" in mime or filename.endswith(".pdf"):
            pdf_att = att
            break

    if not pdf_att:
        return {"error": "Tidak ditemukan attachment PDF pada email ini."}

    pdf_bytes = _download_attachment_bytes(
        service, message_id, pdf_att["attachment_id"]
    )

    parsed = _parse_reimburse_form_pdf(pdf_bytes)
    parsed["source_email_id"] = message_id
    parsed["source_pdf_filename"] = pdf_att.get("filename")

    return parsed

@tool()
def get_email_detail(message_id: str) -> Dict[str, Any]:
    """
    Mengambil detail 1 email, termasuk body dan daftar attachment.

    Args:
        message_id: ID message Gmail (didapat dari list_recent_emails).

    Returns:
        Dict berisi header, body, dan attachments.
    """
    service = _get_gmail_service()

    msg = (
        service.users()
        .messages()
        .get(userId="me", id=message_id, format="full")
        .execute()
    )

    headers = {h["name"]: h["value"] for h in msg["payload"]["headers"]}
    body_text = _extract_text_body(msg.get("payload"))
    attachments = []
    _collect_attachments(msg.get("payload"), attachments)

    return {
        "id": message_id,
        "subject": headers.get("Subject", ""),
        "from": headers.get("From", ""),
        "to": headers.get("To", ""),
        "date": headers.get("Date", ""),
        "snippet": msg.get("snippet", ""),
        "body": body_text,
        "attachments": attachments,
    }

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
