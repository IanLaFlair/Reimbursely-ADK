from typing import List, Dict, Any, Optional
from pathlib import Path

from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from ibm_watsonx_orchestrate.agent_builder.tools import tool
from datetime import date, timedelta
from base64 import b64encode, urlsafe_b64decode

from googleapiclient.errors import HttpError

from io import BytesIO
from PyPDF2 import PdfReader
import re
from pathlib import Path
import requests

import io
import base64
# import pandas as pd


SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]

# BASE_DIR = folder "gmail_tool"
BASE_DIR = Path(__file__).resolve().parents[1]
TOKEN_PATH = BASE_DIR / "token.json"
VISION_KEY_PATH = Path(__file__).resolve().parent.parent / "vision_api_key.txt"


def _get_vision_api_key() -> str | None:
    try:
        return VISION_KEY_PATH.read_text(encoding="utf-8").strip()
    except FileNotFoundError:
        return None


def _get_gmail_service():
    """Load token.json from the tool package and build a Gmail service."""
    creds = Credentials.from_authorized_user_file(str(TOKEN_PATH), SCOPES)
    service = build("gmail", "v1", credentials=creds)
    return service


def _get_current_week_range_until_today():
    """
    Get the date range from Monday of this week until today.

    Returns:
        (monday_str, tomorrow_str) in format YYYY/MM/DD where:
        - monday_str  = Monday of the current week (inclusive, used in `after:`)
        - tomorrow_str = tomorrow (exclusive, used in `before:`)
    """
    today = date.today()
    weekday = today.weekday()  # Monday=0

    monday = today - timedelta(days=weekday)
    tomorrow = today + timedelta(days=1)

    def fmt(d: date) -> str:
        return d.strftime("%Y/%m/%d")

    return fmt(monday), fmt(tomorrow)


VISION_ENDPOINT = "https://vision.googleapis.com/v1/images:annotate"


def _call_vision_ocr(image_bytes: bytes) -> str:
    """
    Call Google Cloud Vision OCR and return the full text result.

    API key is read from tools/gmail_tool/vision_api_key.txt.
    """
    api_key = _get_vision_api_key()
    if not api_key:
        raise RuntimeError(
            "vision_api_key.txt not found or empty. Please put your Cloud Vision API key in that file."
        )

    img_b64 = b64encode(image_bytes).decode("utf-8")

    payload = {
        "requests": [
            {
                "image": {"content": img_b64},
                "features": [{"type": "TEXT_DETECTION"}],
            }
        ]
    }

    params = {"key": api_key}
    resp = requests.post(VISION_ENDPOINT, params=params, json=payload, timeout=30)
    resp.raise_for_status()
    data = resp.json()

    try:
        annotations = data["responses"][0].get("fullTextAnnotation")
        if not annotations:
            return ""
        return annotations.get("text", "")
    except (KeyError, IndexError):
        return ""


def _parse_amounts_from_text(text: str) -> List[int]:
    """
    Extract candidate Rupiah amounts from OCR text.

    Examples it should catch:
      - IDR 300,000
      - Rp 300.000
      - 1.250.000

    Returns:
        List of integer amounts (may be empty).
    """
    if not text:
        return []

    cleaned = text.replace("\n", " ")

    pattern_currency = re.compile(r"(?:IDR|Rp)[^\d]*([0-9\.,]+)", re.IGNORECASE)
    amounts: List[int] = []

    def _to_int(num_str: str) -> int | None:
        s = num_str.strip().replace(",", ".")
        if "." in s:
            parts = s.split(".")
            # if the last part has 2 digits it is likely decimals (e.g. 300.000,00)
            if len(parts[-1]) == 2:
                s = "".join(parts[:-1])
            else:
                s = "".join(parts)
        s = s.replace(".", "")
        try:
            return int(s)
        except ValueError:
            return None

    for m in pattern_currency.findall(cleaned):
        val = _to_int(m)
        if val is not None:
            amounts.append(val)

    if not amounts:
        pattern_generic = re.compile(r"\b(\d{1,3}(?:[.,]\d{3})+)\b")
        for m in pattern_generic.findall(cleaned):
            m_clean = m.replace(".", "").replace(",", "")
            try:
                amounts.append(int(m_clean))
            except ValueError:
                continue

    # unique + sorted
    return sorted(set(amounts))


def _extract_text_body(payload) -> str:
    """Extract the email body (text/plain) from Gmail multipart payload."""
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
    """Download a Gmail attachment as raw bytes."""
    att = (
        service.users()
        .messages()
        .attachments()
        .get(userId="me", messageId=message_id, id=attachment_id)
        .execute()
    )
    data = att.get("data", "")
    return urlsafe_b64decode(data.encode("utf-8"))


def _get_receipt_image_attachments(attachments: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """
    Filter all attachments that look like payment receipts (images).

    Version 1: mimeType image/* or filename ends with .jpg/.jpeg/.png
    """
    receipt_atts: List[Dict[str, Any]] = []
    for att in attachments:
        mime = (att.get("mimeType") or "").lower()
        filename = (att.get("filename") or "").lower()
        if mime.startswith("image/") or filename.endswith((".jpg", ".jpeg", ".png")):
            receipt_atts.append(att)
    return receipt_atts


def _reconcile_form_and_receipts(form_data: Dict[str, Any], receipts: List[Dict[str, Any]]) -> Dict[str, Any]:
    """
    Match items in the reimbursement form with payment receipts.

    form_data:
      {
        "tanggal_pengajuan": "...",
        "items": [
          {"description": "...", "harga": 300000, "jumlah": 1, "subtotal": 300000},
          ...
        ],
        "total": 300000,
        "rekening": {...}
      }

    receipts:
      [
        {
          "filename": "...",
          "candidate_amounts": [...],
          "selected_amount": 300000,
          ...
        },
        ...
      ]
    """
    items: List[Dict[str, Any]] = form_data.get("items", [])
    # use only receipts with a valid selected_amount
    usable_receipts = [
        {"idx": i, "amount": r.get("selected_amount"), "filename": r.get("filename")}
        for i, r in enumerate(receipts)
        if isinstance(r.get("selected_amount"), int)
    ]

    used_flags = {r["idx"]: False for r in usable_receipts}
    per_item_results: List[Dict[str, Any]] = []

    for item in items:
        target = item.get("subtotal")
        matched_idx = None
        matched_filename = None

        if isinstance(target, int):
            for r in usable_receipts:
                if used_flags[r["idx"]]:
                    continue
                if r["amount"] == target:
                    matched_idx = r["idx"]
                    matched_filename = r["filename"]
                    break

        if matched_idx is not None:
            used_flags[matched_idx] = True
            per_item_results.append(
                {
                    "description": item.get("description") or item.get("deskripsi"),
                    "subtotal": target,
                    "status": "MATCH",
                    "receipt_filename": matched_filename,
                }
            )
        else:
            per_item_results.append(
                {
                    "description": item.get("description") or item.get("deskripsi"),
                    "subtotal": target,
                    "status": "MISSING_RECEIPT",
                    "receipt_filename": None,
                }
            )

    # all receipts that were not used
    unmatched_receipts = [
        receipts[idx]
        for idx, used in used_flags.items()
        if not used
    ]

    form_total = form_data.get("total")
    sum_receipts = sum(
        r.get("selected_amount") or 0 for r in receipts if isinstance(r.get("selected_amount"), int)
    )

    overall_status = "OK"
    notes: List[str] = []

    if form_total is not None and isinstance(form_total, int):
        if form_total != sum_receipts:
            overall_status = "MISMATCH"
            notes.append(
                f"Form total ({form_total}) does not match sum of receipt amounts ({sum_receipts})."
            )

    if any(i["status"] != "MATCH" for i in per_item_results):
        overall_status = "MISMATCH"
        notes.append("Some items do not have a matching payment receipt.")

    if unmatched_receipts:
        notes.append("Some payment receipts are unused / do not match any item.")

    return {
        "overall_status": overall_status,
        "items": per_item_results,
        "unmatched_receipts": unmatched_receipts,
        "form_total": form_total,
        "sum_receipt_amounts": sum_receipts,
        "notes": notes,
    }

    # --- unreachable legacy code below (kept as-is, never executed) ---
    items = form["items"]
    receipts_list = receipts["receipts"]

    # mark receipts that are already used
    used = [False] * len(receipts_list)
    hasil_items = []

    for item in items:
        target = item["subtotal"]
        match_idx = None

        for i, rc in enumerate(receipts_list):
            if used[i]:
                continue
            if rc.get("selected_amount") == target:
                match_idx = i
                break

        if match_idx is not None:
            used[match_idx] = True
            hasil_items.append(
                {
                    "description": item["description"],
                    "subtotal": target,
                    "receipt_filename": receipts_list[match_idx]["filename"],
                    "status": "MATCH",
                }
            )
        else:
            hasil_items.append(
                {
                    "description": item["description"],
                    "subtotal": target,
                    "receipt_filename": None,
                    "status": "MISSING_RECEIPT",
                }
            )

    # receipts that remain unused
    unmatched_receipts = [
        rc for i, rc in enumerate(receipts_list) if not used[i]
    ]

    return {
        "items": hasil_items,
        "unmatched_receipts": unmatched_receipts,
    }

def _select_form_pdf_attachment(attachments: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    """
    Choose the most likely PDF that represents the reimbursement *form*,
    not the payment receipt.

    Heuristics:
    - Only consider .pdf attachments.
    - Prefer filenames containing words like 'reimburse', 'form', 'persetujuan', 'approval'.
    - Penalize filenames that look like receipts/invoices (receipt, bukti, invoice, struk).
    """
    pdf_atts: List[Dict[str, Any]] = []
    for att in attachments:
        mime = (att.get("mimeType") or "").lower()
        filename = (att.get("filename") or "").lower()
        if "pdf" in mime or filename.endswith(".pdf"):
            pdf_atts.append(att)

    if not pdf_atts:
        return None

    def score(att: Dict[str, Any]) -> int:
        name = (att.get("filename") or "").lower()
        s = 0

        # Positive signals → likely the PITUKU form
        if any(k in name for k in ["reimburse", "reimbursement", "form", "persetujuan", "approval"]):
            s += 10

        # Negative signals → likely a receipt/invoice instead of the form
        if any(k in name for k in ["receipt", "bukti", "invoice", "struk"]):
            s -= 5

        return s

    pdf_atts.sort(key=score, reverse=True)
    return pdf_atts[0]

def _parse_reimburse_form_pdf(pdf_bytes: bytes) -> Dict[str, Any]:
    """
    Parse the Pituku reimbursement form PDF (like the provided example)
    into structured data.

    Assumes a layout similar to the sample:
      - Tanggal Pengajuan
      - DETAIL PENGAJUAN
      - Nama Bank, Nomor Rekening, Nama Rekening
    """
    reader = PdfReader(BytesIO(pdf_bytes))
    text_parts = []
    for page in reader.pages:
        t = page.extract_text() or ""
        text_parts.append(t)
    full_text = "\n".join(text_parts)

    # Normalize into a list of non-empty lines
    lines = [ln.strip() for ln in full_text.splitlines() if ln.strip()]

    # Tanggal Pengajuan
    tanggal = ""
    for ln in lines:
        if ln.startswith("Tanggal Pengajuan"):
            # e.g. "Tanggal Pengajuan : 16/04/2025"
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

    # Detail pengajuan row
    # We search for the row containing "No Description Detail Harga Jumlah Sub Total"
    item_description = ""
    harga = jumlah = subtotal = None

    # join 3 lines after "No Description Detail Harga Jumlah Sub Total"
    for i, ln in enumerate(lines):
        if ln.startswith("No Description"):
            row_text = " ".join(lines[i + 1: i + 4])  # relatively safe for the sample
            # Description: drop leading "1 " and anything after "Rp"
            m_desc = re.search(r"1\s+(.*?)\s+Rp", row_text)
            if m_desc:
                item_description = m_desc.group(1).strip()

            # Numbers: capture patterns like 300.000, 1.250.000, etc.
            nums = re.findall(r"(\d{1,3}(?:\.\d{3})+)", row_text)

            def to_int(s: str) -> int:
                return int(s.replace(".", ""))

            if nums:
                harga = to_int(nums[0])
                subtotal = to_int(nums[-1])
            # Find quantity (usually 1,2,...) between price & subtotal
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
    """Collect attachments (id, filename, mimeType) from Gmail payload."""
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
def export_reimburse_summary_this_week() -> Dict[str, Any]:
    """
    Generate this week's reimbursement report as an Excel file (Summary sheet only).

    - Fetch all reimbursement emails for this week from Gmail.
    - Analyze each email using analyze_reimburse_email.
    - Build one Summary sheet containing:
        Email ID, Subject, Submission Date, Form Total, Receipt Total, Status, Notes.
    - Return the Excel file as a base64 string.
    """
    import pandas as pd
    service = _get_gmail_service()

    # 1) This week's range (using existing helper)
    start_date, end_date_exclusive = _get_current_week_range_until_today()

    # Same query style as list_reimburse_emails_this_week
    query = f"after:{start_date} before:{end_date_exclusive} reimburse"

    try:
        resp = (
            service.users()
            .messages()
            .list(userId="me", q=query, maxResults=50)
            .execute()
        )
    except HttpError as e:
        return {
            "error": "HttpError while searching reimbursement emails for this week.",
            "status": e.resp.status,
            "reason": e._get_reason(),
            "start_date": start_date,
            "end_date_exclusive": end_date_exclusive,
        }

    messages = resp.get("messages", []) or []

    summary_rows: List[Dict[str, Any]] = []

    for m in messages:
        msg_id = m.get("id")
        if not msg_id:
            continue

        try:
            analysis = analyze_reimburse_email(message_id=msg_id)
        except Exception as e:
            summary_rows.append(
                {
                    "Email ID": msg_id,
                    "Subject": "",
                    "Tanggal Pengajuan": "",
                    "Total Form": None,
                    "Total Bukti Bayar": None,
                    "Status": "ERROR",
                    "Catatan": f"Analysis failed: {e}",
                }
            )
            continue

        form = analysis.get("form_data", {}) or {}
        recon = analysis.get("reconciliation", {}) or {}

        tanggal_pengajuan = form.get("tanggal_pengajuan", "")
        total_form = form.get("total")
        total_receipts = recon.get("sum_receipt_amounts")

        status = recon.get("overall_status", "UNKNOWN")
        notes_list = recon.get("notes", []) or []
        catatan = "; ".join(notes_list)

        summary_rows.append(
            {
                "Email ID": msg_id,
                "Subject": analysis.get("subject", ""),
                "Tanggal Pengajuan": tanggal_pengajuan,
                "Total Form": total_form,
                "Total Bukti Bayar": total_receipts,
                "Status": status,
                "Catatan": catatan,
            }
        )

    # 2) If there are no emails, still return info
    if not summary_rows:
        return {
            "filename": "reimbursely_summary_minggu_ini.xlsx",
            "mime_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "file_base64": None,
            "row_count": 0,
            "start_date": start_date,
            "end_date_exclusive": end_date_exclusive,
            "info": "No reimbursement emails found for this week.",
        }

    # 3) Build DataFrame and Excel in memory
    df = pd.DataFrame(summary_rows)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Summary", index=False)

        # Basic formatting: bold header, autofilter, simple auto width
        workbook = writer.book
        worksheet = writer.sheets["Summary"]

        header_format = workbook.add_format({"bold": True})
        worksheet.set_row(0, None, header_format)

        # Autofilter entire range
        worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)

        # Auto column width
        for col_idx, col_name in enumerate(df.columns):
            max_len = max(
                df[col_name].astype(str).map(len).max(),
                len(col_name),
            )
            worksheet.set_column(col_idx, col_idx, max_len + 2)

    excel_bytes = buffer.getvalue()
    file_b64 = base64.b64encode(excel_bytes).decode("utf-8")

    return {
        "filename": "reimbursely_summary_minggu_ini.xlsx",
        "mime_type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "file_base64": file_b64,
        "row_count": len(summary_rows),
        "start_date": start_date,
        "end_date_exclusive": end_date_exclusive,
    }


@tool()
def extract_all_payment_amounts_from_email(message_id: str) -> Dict[str, Any]:
    """
    OCR ALL payment receipt images in this email and extract payment amounts
    from each attachment.

    Useful when one form has multiple items and each item has a receipt.
    """
    service = _get_gmail_service()

    try:
        msg = (
            service.users()
            .messages()
            .get(userId="me", id=message_id, format="full")
            .execute()
        )
    except HttpError as e:
        return {
            "error": "HttpError while fetching message for payment receipts.",
            "status": e.resp.status,
            "reason": e._get_reason(),
            "message_id": message_id,
        }

    attachments = []
    _collect_attachments(msg.get("payload"), attachments)

    receipt_atts = _get_receipt_image_attachments(attachments)
    if not receipt_atts:
        return {
            "error": "No image attachments (payment receipts) found in this email.",
            "message_id": message_id,
            "attachments_ditemukan": attachments,
        }

    receipts_result = []

    for att in receipt_atts:
        try:
            img_bytes = _download_attachment_bytes(
                service, message_id, att["attachment_id"]
            )
        except HttpError as e:
            receipts_result.append(
                {
                    "filename": att.get("filename"),
                    "error": "Failed to download attachment.",
                    "status": e.resp.status,
                    "reason": e._get_reason(),
                }
            )
            continue

        try:
            ocr_text = _call_vision_ocr(img_bytes)
            amounts = _parse_amounts_from_text(ocr_text)
            selected = amounts[-1] if amounts else None
        except Exception as e:
            receipts_result.append(
                {
                    "filename": att.get("filename"),
                    "error": f"OCR failed: {e}",
                }
            )
            continue

        receipts_result.append(
            {
                "filename": att.get("filename"),
                "candidate_amounts": amounts,
                "selected_amount": selected,
                "ocr_text_preview": (ocr_text[:300] + "…") if ocr_text else "",
            }
        )

    return {
        "message_id": message_id,
        "receipts": receipts_result,
    }


@tool()
def list_reimburse_emails_this_week(max_results: int = 50) -> Dict[str, Any]:
    """
    Fetch reimbursement emails for the period Monday–today.

    Intended for testing or weekly processing.
    """
    start_date, tomorrow = _get_current_week_range_until_today()
    return list_reimburse_emails_for_period(start_date, tomorrow, max_results)


@tool()
def analyze_reimburse_email(message_id: str) -> Dict[str, Any]:
    """
    Fully analyze one reimbursement email:

    - Parse reimbursement form (PDF) from attachments.
    - OCR all payment receipts (image attachments).
    - Reconcile form items against receipts.
    - Return status OK / MISMATCH and detailed breakdown.

    Args:
        message_id: Gmail message ID.

    Returns:
        Dict containing:
          - form_data
          - receipts
          - reconciliation (overall_status, per item, unmatched, totals)
        or an error dict if something fails.
    """
    service = _get_gmail_service()

    # 1) Fetch the email
    try:
        msg = (
            service.users()
            .messages()
            .get(userId="me", id=message_id, format="full")
            .execute()
        )

        # Extract subject from headers
        payload = msg.get("payload", {})
        headers = payload.get("headers", [])

        subject = ""
        for h in headers:
            name = h.get("name", "").lower()
            if name == "subject":
                subject = h.get("value", "")
                break

    except HttpError as e:
        return {
            "error": "HttpError while fetching message.",
            "status": e.resp.status,
            "reason": e._get_reason(),
            "message_id": message_id,
        }

    attachments: List[Dict[str, Any]] = []
    _collect_attachments(msg.get("payload"), attachments)

    if not attachments:
        return {
            "error": "No attachments found in this email.",
            "message_id": message_id,
        }

    # 2) Find reimbursement PDF form (first PDF attachment)
    form_att = _select_form_pdf_attachment(attachments)

    if not form_att:
        return {
            "error": "No reimbursement form PDF attachment found in this email.",
            "message_id": message_id,
            "attachments_ditemukan": attachments,
        }

    # Download & parse PDF form
    try:
        pdf_bytes = _download_attachment_bytes(
            service, message_id, form_att["attachment_id"]
        )
    except HttpError as e:
        return {
            "error": "HttpError while downloading PDF form.",
            "status": e.resp.status,
            "reason": e._get_reason(),
            "message_id": message_id,
            "attachment_id": form_att.get("attachment_id"),
            "filename": form_att.get("filename"),
        }

    try:
        form_data = _parse_reimburse_form_pdf(pdf_bytes)
    except Exception as e:
        return {
            "error": f"Failed to parse reimbursement PDF form: {e}",
            "message_id": message_id,
            "filename": form_att.get("filename"),
        }

    form_data["source_email_id"] = message_id
    form_data["source_pdf_filename"] = form_att.get("filename")

    # 3) OCR all payment receipt images
    receipt_atts = _get_receipt_image_attachments(attachments)
    receipts_result: List[Dict[str, Any]] = []

    for att in receipt_atts:
        filename = att.get("filename")
        try:
            img_bytes = _download_attachment_bytes(
                service, message_id, att["attachment_id"]
            )
        except HttpError as e:
            receipts_result.append(
                {
                    "filename": filename,
                    "error": "Failed to download payment receipt attachment.",
                    "status": e.resp.status,
                    "reason": e._get_reason(),
                }
            )
            continue

        try:
            ocr_text = _call_vision_ocr(img_bytes)
            amounts = _parse_amounts_from_text(ocr_text)
            selected = amounts[-1] if amounts else None
        except Exception as e:
            receipts_result.append(
                {
                    "filename": filename,
                    "error": f"OCR failed: {e}",
                }
            )
            continue

        receipts_result.append(
            {
                "filename": filename,
                "candidate_amounts": amounts,
                "selected_amount": selected,
                "ocr_text_preview": (ocr_text[:300] + "…") if ocr_text else "",
            }
        )

    # 4) Reconcile form vs receipts
    reconciliation = _reconcile_form_and_receipts(form_data, receipts_result)

    return {
        "message_id": message_id,
        "subject": subject,
        "form_data": form_data,
        "receipts": receipts_result,
        "reconciliation": reconciliation,
    }


@tool()
def list_reimburse_emails_for_period(start_date: str, end_date: str, max_results: int = 50) -> Dict[str, Any]:
    """
    Fetch reimbursement emails within a specific date range.

    Args:
        start_date: Start date (inclusive) in YYYY/MM/DD format, e.g. '2025/04/14'.
        end_date: End date (exclusive) in YYYY/MM/DD format, used in 'before:'.
        max_results: Max number of emails to fetch.

    Returns:
        Dict containing period info and list of emails (id, subject, from, date).
    """
    service = _get_gmail_service()
    finance_addr = "fak3optic@gmail.com"

    query = (
        f'("reimburse" OR "reimbursement") '
        f'-subject:advance '
        f'(to:{finance_addr} OR cc:{finance_addr}) '
        f'after:{start_date} before:{end_date}'
    )

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
        subject = headers.get("Subject", "") or ""

        # Skip cash advance / advance fund requests
        if "advance" in subject.lower():
            continue

        emails.append(
            {
                "id": m["id"],
                "subject": subject,
                "from": headers.get("From", ""),
                "date": headers.get("Date", ""),
                "internal_ts": int(msg.get("internalDate", 0)),
            }
        )

    # sort newest-first by internalDate
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
def extract_all_payment_amounts_from_email(message_id: str) -> Dict[str, Any]:
    """
    OCR ALL payment receipt images in this email and extract payment amounts
    from each attachment.

    Args:
        message_id: Gmail message ID.

    Returns:
        {
          "message_id": "...",
          "receipts": [
            {
              "filename": "...",
              "candidate_amounts": [300000],
              "selected_amount": 300000,
              "ocr_text_preview": "Top Up Success IDR 300,000 ..."
            },
            ...
          ]
        }
        or an error dict.
    """
    service = _get_gmail_service()

    try:
        msg = (
            service.users()
            .messages()
            .get(userId="me", id=message_id, format="full")
            .execute()
        )
    except HttpError as e:
        return {
            "error": "HttpError while fetching message for payment receipts.",
            "status": e.resp.status,
            "reason": e._get_reason(),
            "message_id": message_id,
        }

    attachments: List[Dict[str, Any]] = []
    _collect_attachments(msg.get("payload"), attachments)

    receipt_atts = _get_receipt_image_attachments(attachments)
    if not receipt_atts:
        return {
            "error": "No image attachments (payment receipts) found in this email.",
            "message_id": message_id,
            "attachments_ditemukan": attachments,
        }

    receipts_result: List[Dict[str, Any]] = []

    for att in receipt_atts:
        filename = att.get("filename")
        try:
            img_bytes = _download_attachment_bytes(
                service, message_id, att["attachment_id"]
            )
        except HttpError as e:
            receipts_result.append(
                {
                    "filename": filename,
                    "error": "Failed to download attachment.",
                    "status": e.resp.status,
                    "reason": e._get_reason(),
                }
            )
            continue

        try:
            ocr_text = _call_vision_ocr(img_bytes)
            amounts = _parse_amounts_from_text(ocr_text)
            selected = amounts[-1] if amounts else None  # use the largest/last
        except Exception as e:
            receipts_result.append(
                {
                    "filename": filename,
                    "error": f"OCR failed: {e}",
                }
            )
            continue

        receipts_result.append(
            {
                "filename": filename,
                "candidate_amounts": amounts,
                "selected_amount": selected,
                "ocr_text_preview": (ocr_text[:300] + "…") if ocr_text else "",
            }
        )

    return {
        "message_id": message_id,
        "receipts": receipts_result,
    }


@tool()
def parse_reimburse_form_from_email(message_id: str) -> Dict[str, Any]:
    """
    Find the first PDF attachment in this email, download it, and parse
    the reimbursement form.

    Args:
        message_id: Gmail message ID.

    Returns:
        Dict containing reimbursement summary (date, items, total, bank info),
        or an error dict if Gmail API / parsing fails.
    """
    service = _get_gmail_service()

    try:
        msg = (
            service.users()
            .messages()
            .get(userId="me", id=message_id, format="full")
            .execute()
        )
    except HttpError as e:
        return {
            "error": "HttpError while fetching message.",
            "status": e.resp.status,
            "reason": e._get_reason(),
            "message_id": message_id,
        }

    attachments = []
    _collect_attachments(msg.get("payload"), attachments)

    # pick the first PDF attachment
    pdf_att = _select_form_pdf_attachment(attachments)

    if not pdf_att:
        return {
            "error": "No PDF attachment found in this email.",
            "message_id": message_id,
            "attachments_ditemukan": attachments,
        }

    try:
        pdf_bytes = _download_attachment_bytes(
            service, message_id, pdf_att["attachment_id"]
        )
    except HttpError as e:
        return {
            "error": "HttpError while downloading attachment.",
            "status": e.resp.status,
            "reason": e._get_reason(),
            "message_id": message_id,
            "attachment_id": pdf_att.get("attachment_id"),
            "filename": pdf_att.get("filename"),
        }

    # If we get here, download was fine; now parse the PDF.
    try:
        parsed = _parse_reimburse_form_pdf(pdf_bytes)
    except Exception as e:
        return {
            "error": f"Failed to parse PDF: {e}",
            "message_id": message_id,
            "filename": pdf_att.get("filename"),
        }

    parsed["source_email_id"] = message_id
    parsed["source_pdf_filename"] = pdf_att.get("filename")

    return parsed


@tool()
def get_email_detail(message_id: str) -> Dict[str, Any]:
    """
    Fetch detailed information of a single email, including body and
    attachments list.

    Args:
        message_id: Gmail message ID (from list_recent_emails).

    Returns:
        Dict containing header fields, body, and attachments.
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
    Fetch a list of recent emails from Gmail.

    Args:
        query: Gmail query string (e.g. 'subject:reimbursement OR subject:claim').
        max_results: Max number of emails to return.

    Returns:
        List of dicts containing id, subject, from, date, snippet.
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
