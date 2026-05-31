import base64
import json
import logging
import os
import re
import zipfile
from io import BytesIO

import requests
from dotenv import load_dotenv
from flask import Flask, jsonify, request, send_file, send_from_directory
from flask_cors import CORS
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)-8s %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("factory-scanner")

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
PUBLIC_DIR = os.path.join(BASE_DIR, "public")
app = Flask(__name__, static_folder=PUBLIC_DIR, static_url_path="/_static")
CORS(app, resources={r"/api/*": {"origins": "*"}, r"/*": {"origins": "*"}})

app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024

GROQ_API_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_MODEL   = os.environ.get("GROQ_MODEL", "meta-llama/llama-4-scout-17b-16e-instruct")
GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "").strip()

if not GROQ_API_KEY:
    log.warning("GROQ_API_KEY not set — /extract will return 500.")


# ── OCR PROMPT ────────────────────────────────────────────────────────────────

OCR_PROMPT = """\
Extract ALL content from this image into a JSON table preserving the exact grid layout.
Output ONLY this JSON, nothing else:
{"table":[["તારીખ","35","36",...],["1",34,null,...],["2",null,35,...],...]}

RULES:
1. Include EVERY row and column visible — headers, labels, data, blank rows — nothing skipped.
2. Numbers → JSON number (34, not "34"). Text (including Gujarati ગુજરાતી script) → JSON string exactly as written. Blank/empty/illegible → null.
3. Gujarati and other non-Latin scripts: copy the exact Unicode characters as-is into the JSON string.
4. All rows must have the same number of elements (pad with null on the right).
5. Handwriting: 1↔7, 0↔6, 3↔8, 4↔9 are common confusions — read carefully.
6. Zero (0) is valid, never replace with null.
7. Output ONLY the raw JSON object, no markdown, no explanation.\
"""


# ── IMAGE PREPROCESSING ───────────────────────────────────────────────────────

def preprocess_image(image_b64: str) -> tuple[str, str]:
    try:
        from PIL import Image, ImageEnhance
        raw = base64.b64decode(image_b64)
        img = Image.open(BytesIO(raw)).convert("RGB")
        max_dim = max(img.width, img.height)
        if max_dim < 1400:
            scale = 1400 / max_dim
            img = img.resize(
                (min(int(img.width * scale), 2800), min(int(img.height * scale), 2800)),
                Image.LANCZOS,
            )
        elif max_dim > 2800:
            scale = 2800 / max_dim
            img = img.resize((int(img.width * scale), int(img.height * scale)), Image.LANCZOS)
        img = ImageEnhance.Contrast(img).enhance(1.6)
        img = ImageEnhance.Sharpness(img).enhance(2.8)
        img = ImageEnhance.Brightness(img).enhance(1.08)
        buf = BytesIO()
        img.save(buf, format="JPEG", quality=92, optimize=True)
        buf.seek(0)
        return base64.b64encode(buf.getvalue()).decode(), "image/jpeg"
    except Exception as exc:
        log.warning("Preprocessing failed (%s); using original.", exc)
        return image_b64, "image/jpeg"


# ── GROQ API ──────────────────────────────────────────────────────────────────

def call_groq(image_b64: str, media_type: str) -> requests.Response:
    payload = {
        "model": GROQ_MODEL,
        "messages": [{
            "role": "user",
            "content": [
                {"type": "text", "text": OCR_PROMPT},
                {"type": "image_url", "image_url": {"url": f"data:{media_type};base64,{image_b64}"}},
            ],
        }],
        "temperature": 0,
        "max_completion_tokens": 8192,
        "response_format": {"type": "json_object"},
    }
    headers = {"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"}
    return requests.post(GROQ_API_URL, headers=headers, json=payload, timeout=120)


# ── URL IMAGE FETCHING ────────────────────────────────────────────────────────

def fetch_image_from_url(url: str) -> tuple[str, str]:
    import ipaddress
    from urllib.parse import urlparse

    parsed = urlparse(url)
    if parsed.scheme not in ("http", "https"):
        raise ValueError("Only http/https URLs are supported.")
    if not parsed.hostname:
        raise ValueError("Invalid URL.")

    # Basic SSRF protection: block private/loopback/reserved IPs
    try:
        addr = ipaddress.ip_address(parsed.hostname)
        if addr.is_private or addr.is_loopback or addr.is_reserved:
            raise ValueError("Private or reserved IP addresses are not allowed.")
    except ValueError as exc:
        if "Private" in str(exc) or "reserved" in str(exc) or "loopback" in str(exc):
            raise

    resp = requests.get(url, timeout=15, stream=True)
    resp.raise_for_status()

    content_type = resp.headers.get("content-type", "image/jpeg").split(";")[0].strip()
    if not content_type.startswith("image/"):
        raise ValueError(f"URL does not point to an image (got: {content_type}).")

    data = resp.content
    if len(data) > 10 * 1024 * 1024:
        raise ValueError("Image from URL exceeds the 10 MB limit.")

    return base64.b64encode(data).decode(), content_type


# ── HELPERS ───────────────────────────────────────────────────────────────────

def extract_json(text: str) -> dict:
    cleaned = text.strip().replace("```json", "").replace("```", "").strip()
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", cleaned, re.DOTALL)
        if match:
            return json.loads(match.group(0))
        raise


def _coerce_to_table(parsed: dict) -> list | None:
    """
    Try every key/format the model might return and produce a 2D list.
    Handles: {table:[...]}, {rows:[{label,values}]}, {data:[...]}, {sheet:[...]},
             {rows:[[...]]}, top-level list, etc.
    """
    # Direct 2D array keys
    for key in ("table", "data", "sheet", "cells", "matrix", "grid"):
        val = parsed.get(key)
        if isinstance(val, list) and val:
            # Already 2D?
            if isinstance(val[0], list):
                return val
            # List of dicts (old {label, values} format)?
            if isinstance(val[0], dict) and "values" in val[0]:
                rows = []
                for r in val:
                    label  = r.get("label", "")
                    values = r.get("values", [])
                    rows.append([label] + list(values))
                return rows or None

    # rows key — could be list-of-dicts or list-of-lists
    rows_val = parsed.get("rows")
    if isinstance(rows_val, list) and rows_val:
        if isinstance(rows_val[0], list):
            return rows_val
        if isinstance(rows_val[0], dict):
            rows = []
            for r in rows_val:
                label  = r.get("label", "")
                values = r.get("values", [])
                rows.append([label] + list(values))
            return rows or None

    return None


def normalize_cell(v):
    """Normalize a cell value to int, float, str, or None. Preserves all text including Gujarati."""
    if v is None:
        return None
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, int):
        return v
    if isinstance(v, float):
        return int(v) if v == int(v) else v
    s = str(v).strip()
    if not s:
        return None
    # Try numeric conversion first
    try:
        return int(s)
    except ValueError:
        pass
    try:
        f = float(s)
        return int(f) if f == int(f) else f
    except ValueError:
        pass
    return s  # preserve text as-is (Gujarati, English, etc.)


def normalize_table(raw) -> list:
    """Validate and normalize a 2D table."""
    if not raw or not isinstance(raw, list):
        return []
    rows = []
    for row in raw:
        if isinstance(row, list):
            rows.append([normalize_cell(v) for v in row])
        elif isinstance(row, dict):
            # Gracefully handle a row that came back as a dict
            rows.append([normalize_cell(v) for v in row.values()])
        else:
            rows.append([normalize_cell(row)])
    if not rows:
        return []
    num_cols = max((len(r) for r in rows), default=0)
    if not num_cols:
        return []
    for row in rows:
        while len(row) < num_cols:
            row.append(None)
    return rows


def safe_filename(value, fallback: str = "factory-data") -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "-", str(value or fallback)).strip("-._")
    return cleaned or fallback


def err(message: str, status: int = 400):
    log.warning("Response %d: %s", status, message)
    return jsonify({"error": message}), status


# ── EXCEL BUILDER ─────────────────────────────────────────────────────────────

def build_excel(table) -> BytesIO | None:
    table = normalize_table(table)
    if not table:
        return None

    num_rows = len(table)
    num_cols = len(table[0])

    wb = Workbook()
    ws = wb.active
    ws.title = "Extracted Data"

    header_fill  = PatternFill("solid", fgColor="D9E1F2")
    header_font  = Font(bold=True)
    total_fill   = PatternFill("solid", fgColor="E2EFDA")
    grand_fill   = PatternFill("solid", fgColor="C6EFCE")
    total_font   = Font(bold=True)
    center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
    wrap_align   = Alignment(wrap_text=True, vertical="top")

    total_col = num_cols + 1  # Excel column index for row-total column

    # Write extracted table + row-total formula per data row
    for ri, row in enumerate(table, start=1):
        is_hdr = (ri == 1)
        for ci, val in enumerate(row, start=1):
            cell = ws.cell(row=ri, column=ci, value=val)
            if is_hdr:
                cell.font      = header_font
                cell.fill      = header_fill
                cell.alignment = center_align
            else:
                cell.alignment = wrap_align

        if is_hdr:
            # "Total" header in the extra total column
            hc = ws.cell(row=1, column=total_col, value="Total")
            hc.font = header_font; hc.fill = header_fill; hc.alignment = center_align
        elif num_cols > 1:
            # Row total: sum cols B..last_data (skip col A = labels)
            first = get_column_letter(2)
            last  = get_column_letter(num_cols)
            tc = ws.cell(row=ri, column=total_col,
                         value=f"=SUM({first}{ri}:{last}{ri})")
            tc.font = total_font; tc.fill = total_fill; tc.alignment = center_align

    # Column-total row at the bottom (skip row 1 = header, skip col 1 = labels)
    tr = num_rows + 1
    lbl = ws.cell(row=tr, column=1, value="Total")
    lbl.font = total_font; lbl.fill = total_fill; lbl.alignment = center_align

    for ci in range(2, num_cols + 1):
        cl = get_column_letter(ci)
        tc = ws.cell(row=tr, column=ci, value=f"=SUM({cl}2:{cl}{num_rows})")
        tc.font = total_font; tc.fill = total_fill; tc.alignment = center_align

    # Grand total (sum of total column, rows 2..num_rows)
    gtcl = get_column_letter(total_col)
    gc = ws.cell(row=tr, column=total_col,
                 value=f"=SUM({gtcl}2:{gtcl}{num_rows})")
    gc.font = Font(bold=True); gc.fill = grand_fill; gc.alignment = center_align

    # Auto column widths
    all_rows = table + [["Total"] + [None] * num_cols]
    for ci in range(1, total_col + 1):
        max_len = 0
        for row in all_rows:
            if ci <= len(row):
                v = row[ci - 1]
                if v is not None:
                    max_len = max(max_len, len(str(v)))
        ws.column_dimensions[get_column_letter(ci)].width = min(max(max_len + 2, 8), 50)

    ws.freeze_panes = "B2"  # freeze header row + label column

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ── SECURITY HEADERS ──────────────────────────────────────────────────────────

@app.after_request
def add_security_headers(response):
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"]         = "DENY"
    response.headers["Referrer-Policy"]          = "strict-origin-when-cross-origin"
    return response


# ── ROUTES ────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return send_from_directory(PUBLIC_DIR, "index.html")


@app.route("/health")
def health():
    return jsonify({"ok": True, "model": GROQ_MODEL, "provider": "groq"})


@app.route("/extract", methods=["POST"])
def extract():
    if not GROQ_API_KEY:
        return err("GROQ_API_KEY is not configured on the server.", 500)

    data       = request.get_json(silent=True) or {}
    image_b64  = data.get("image_base64")
    media_type = data.get("media_type", "image/jpeg")
    image_url  = data.get("image_url")

    if image_url and not image_b64:
        try:
            image_b64, media_type = fetch_image_from_url(image_url)
        except Exception as exc:
            return err(f"Could not load image from URL: {exc}", 400)

    if not image_b64:
        return err("No image provided.", 400)
    if not isinstance(image_b64, str):
        return err("image_base64 must be a string.", 400)
    if len(image_b64.encode()) > 6 * 1024 * 1024:
        return err("Image payload too large. Reduce photo size or resolution.", 413)

    log.info("Extracting: media=%s url=%s", media_type, bool(image_url))

    processed_b64, processed_type = preprocess_image(image_b64)

    for attempt in range(2):
        try:
            resp = call_groq(processed_b64, processed_type)
        except requests.Timeout:
            if attempt == 0:
                log.warning("Groq timeout on attempt 1; retrying.")
                continue
            return err("Request to Groq timed out. Please try again.", 504)
        except requests.RequestException as exc:
            return err(f"Could not reach Groq API: {exc}", 502)

        if resp.status_code != 200:
            log.error("Groq error %d: %s", resp.status_code, resp.text[:400])
            if attempt == 0:
                continue
            return err(f"Groq API returned error {resp.status_code}. Please try again.", 502)

        try:
            choice  = resp.json()["choices"][0]
            text    = choice["message"]["content"]
            finish  = choice.get("finish_reason", "")
            if finish == "length":
                log.warning("Model hit token limit (attempt %d) — response truncated.", attempt + 1)
            parsed = extract_json(text)
        except Exception as exc:
            log.error("JSON parse error (attempt %d): %s | raw: %.300s", attempt + 1, exc, text if 'text' in dir() else '')
            if attempt == 0:
                continue
            return err("Could not parse the AI response. Try a clearer image.", 500)

        raw_table = _coerce_to_table(parsed)
        if raw_table is None:
            log.warning("No recognized table key in response (attempt %d). Keys: %s", attempt + 1, list(parsed.keys()))
            if attempt == 0:
                continue
            return err(
                "No structured data found in the image. "
                "Make sure the image contains a table, spreadsheet, or form.",
                422,
            )

        table = normalize_table(raw_table)
        if not table:
            if attempt == 0:
                log.warning("Table normalized to empty; retrying.")
                continue
            return err(
                "No structured data found in the image. "
                "Make sure the image contains a table, spreadsheet, or form.",
                422,
            )

        log.info(
            "Extracted table: %d rows × %d cols.",
            len(table),
            len(table[0]) if table else 0,
        )
        return jsonify({"table": table})

    return err("Extraction failed after multiple attempts. Please use a clearer image.", 500)


@app.route("/download-excel", methods=["POST"])
def download_excel():
    data   = request.get_json(silent=True) or {}
    output = build_excel(data.get("table"))
    if output is None:
        return err("No table data provided.", 400)

    filename = safe_filename(data.get("filename"), "factory-data")
    log.info("Serving Excel: %s.xlsx", filename)
    return send_file(
        output,
        as_attachment=True,
        download_name=f"{filename}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/download-all-excels", methods=["POST"])
def download_all_excels():
    data      = request.get_json(silent=True) or {}
    documents = data.get("documents") or []
    ready     = [d for d in documents if d.get("table")]
    if not ready:
        return err("No completed documents provided.", 400)

    buf  = BytesIO()
    used: set[str] = set()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for i, doc in enumerate(ready, start=1):
            wb = build_excel(doc.get("table"))
            if wb is None:
                continue
            base  = safe_filename(doc.get("filename"), f"image-{i}")
            fname = f"{base}.xlsx"
            n = 2
            while fname.lower() in used:
                fname = f"{base}-{n}.xlsx"
                n += 1
            used.add(fname.lower())
            zf.writestr(fname, wb.getvalue())

    buf.seek(0)
    log.info("Serving ZIP with %d Excel file(s).", len(ready))
    return send_file(
        buf,
        as_attachment=True,
        download_name="factory-scanner-excels.zip",
        mimetype="application/zip",
    )


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    log.info("Starting Factory Scanner (dev) on http://localhost:%d", port)
    app.run(host="0.0.0.0", debug=False, port=port)
