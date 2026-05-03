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
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

load_dotenv()

# ── Logging ───────────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)-8s %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
log = logging.getLogger("factory-scanner")

# ── App ───────────────────────────────────────────────────────────────────────
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PUBLIC_DIR = os.path.join(BASE_DIR, "public")
app = Flask(__name__, static_folder=PUBLIC_DIR, static_url_path="/_static")
CORS(app, resources={r"/api/*": {"origins": "*"}, r"/*": {"origins": "*"}})

app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024  # 20 MB hard limit

# ── Config ────────────────────────────────────────────────────────────────────
GROQ_API_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_MODEL   = os.environ.get("GROQ_MODEL", "meta-llama/llama-4-scout-17b-16e-instruct")
GROQ_API_KEY = os.environ.get("GROQ_API_KEY", "").strip()

if not GROQ_API_KEY:
    log.warning("GROQ_API_KEY is not set — /extract will return 500 until configured.")

OCR_PROMPT = """\
You are an expert at reading handwritten factory production records, tally sheets, and ledger books.

Your task: Carefully examine this image and extract ALL numeric data into a structured table that matches the sheet's row and column layout.

Return ONLY valid JSON in this exact format (no markdown, no explanation, no code fences):
{
  "rows": [
    {"label": "1", "values": [34, 32, 35, null]},
    {"label": "", "values": [null, null, null, null]},
    {"label": "3", "values": [35, 34, 37, 28]}
  ]
}

EXTRACTION RULES:
1. Preserve the full data grid top-to-bottom: include EVERY body row that belongs in the table, in visual order — do not skip or merge rows. If a row is blank, only partially filled, or clearly cancelled/struck-through, still emit one JSON object for it with null for every empty or cancelled cell (never omit that row).
2. "label" = the leftmost column text if present (row number, name, date, shift, etc.). If that cell is blank, use an empty string "" for "label". Do not renumber or collapse rows to hide blanks.
3. "values" = numeric cells only, strictly left-to-right, one entry per printed data column — including columns that are intentionally blank on the sheet (use null for those positions so column alignment matches the image).
4. Read handwritten digits with extra care — common confusions: 1/7, 0/6, 3/8, 4/9, 5/6.
5. Blank, smudged, crossed-out, or truly unreadable cells → null (never guess; never substitute 0 for blank). Cancelled or voided rows still appear as a row of nulls; use "" for "label" when the label cell is blank or illegible.
6. EXCLUDE: column header row, any pre-written totals/grand-total rows, date-only header rows.
7. INCLUDE: all actual data values including legitimate zeros (0).
8. All rows MUST have the same number of values; pad shorter rows with null on the right until they match the widest row.
9. If the table has multiple sections separated by sub-headers, include all data rows in one flat list in reading order; blank spacer rows between sections still count as rows (all null values).
10. Return ONLY the raw JSON object — nothing else.\
"""


# ── Image preprocessing ───────────────────────────────────────────────────────
def preprocess_image(image_b64: str) -> tuple[str, str]:
    """Enhance contrast and sharpness for better OCR accuracy."""
    try:
        from PIL import Image, ImageEnhance

        raw = base64.b64decode(image_b64)
        img = Image.open(BytesIO(raw)).convert("RGB")

        # Normalise resolution
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
        log.warning("Image preprocessing failed (%s); using original.", exc)
        return image_b64, "image/jpeg"


# ── Groq call ─────────────────────────────────────────────────────────────────
def call_groq(image_b64: str, media_type: str) -> requests.Response:
    payload = {
        "model": GROQ_MODEL,
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": OCR_PROMPT},
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:{media_type};base64,{image_b64}"},
                    },
                ],
            }
        ],
        "temperature": 0,
        "max_completion_tokens": 8192,
        "response_format": {"type": "json_object"},
    }
    headers = {"Authorization": f"Bearer {GROQ_API_KEY}", "Content-Type": "application/json"}
    return requests.post(GROQ_API_URL, headers=headers, json=payload, timeout=120)


# ── Helpers ───────────────────────────────────────────────────────────────────
def extract_json(text: str) -> dict:
    cleaned = text.strip().replace("```json", "").replace("```", "").strip()
    try:
        return json.loads(cleaned)
    except json.JSONDecodeError:
        match = re.search(r"\{.*\}", cleaned, re.DOTALL)
        if match:
            return json.loads(match.group(0))
        raise


def normalize_rows(rows: list | None) -> list:
    out = []
    for i, row in enumerate(rows or [], start=1):
        values = row.get("values", []) if isinstance(row, dict) else []
        if isinstance(row, dict) and "label" in row:
            raw_lbl = row["label"]
            label = "" if raw_lbl is None else str(raw_lbl)
        else:
            label = str(i)
        out.append(
            {
                "label": label,
                "values": [v if v not in ("", "null") else None for v in values],
            }
        )
    return out


def as_number(value):
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def safe_filename(value, fallback: str = "factory-data") -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "-", str(value or fallback)).strip("-._")
    return cleaned or fallback


def err(message: str, status: int = 400):
    log.warning("Response %d: %s", status, message)
    return jsonify({"error": message}), status


# ── Excel builder ─────────────────────────────────────────────────────────────
def build_excel(rows: list) -> BytesIO | None:
    rows = normalize_rows(rows)
    if not rows:
        return None

    num_cols = max(1, *(len(r["values"]) for r in rows))
    for row in rows:
        while len(row["values"]) < num_cols:
            row["values"].append(None)

    wb = Workbook()
    ws = wb.active
    ws.title = "Factory Data"

    headers = ["Row", *[f"Col {i}" for i in range(1, num_cols + 1)], "Row Total"]
    ws.append(headers)

    for ri, row in enumerate(rows, start=2):
        ws.cell(row=ri, column=1, value=row["label"])
        for ci, val in enumerate(row["values"], start=2):
            ws.cell(row=ri, column=ci, value=as_number(val))
        first_col = get_column_letter(2)
        last_col  = get_column_letter(num_cols + 1)
        ws.cell(row=ri, column=num_cols + 2, value=f"=SUM({first_col}{ri}:{last_col}{ri})")

    total_row = len(rows) + 2
    ws.cell(row=total_row, column=1, value="Grand Total")
    for ci in range(2, num_cols + 2):
        cl = get_column_letter(ci)
        ws.cell(row=total_row, column=ci, value=f"=SUM({cl}2:{cl}{total_row - 1})")
    tcl = get_column_letter(num_cols + 2)
    ws.cell(row=total_row, column=num_cols + 2, value=f"=SUM({tcl}2:{tcl}{total_row - 1})")

    header_fill = PatternFill("solid", fgColor="E8EEF7")
    total_fill  = PatternFill("solid", fgColor="EAF7EF")
    for cell in ws[1]:
        cell.font = Font(bold=True); cell.fill = header_fill
    for cell in ws[total_row]:
        cell.font = Font(bold=True); cell.fill = total_fill

    ws.freeze_panes = "B2"
    ws.auto_filter.ref = ws.dimensions
    ws.column_dimensions["A"].width = 14
    for ci in range(2, num_cols + 2):
        ws.column_dimensions[get_column_letter(ci)].width = 10
    ws.column_dimensions[get_column_letter(num_cols + 2)].width = 13

    buf = BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ── Security headers ──────────────────────────────────────────────────────────
@app.after_request
def add_security_headers(response):
    response.headers["X-Content-Type-Options"] = "nosniff"
    response.headers["X-Frame-Options"]         = "DENY"
    response.headers["Referrer-Policy"]          = "strict-origin-when-cross-origin"
    return response


# ── Routes ────────────────────────────────────────────────────────────────────
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

    data = request.get_json(silent=True) or {}
    image_b64  = data.get("image_base64")
    media_type = data.get("media_type", "image/jpeg")

    if not image_b64:
        return err("No image provided.", 400)
    if not isinstance(image_b64, str):
        return err("image_base64 must be a string.", 400)
    if len(image_b64.encode()) > 6 * 1024 * 1024:
        return err("Image payload too large. Reduce photo size or resolution.", 413)

    log.info("Extracting: %.60s…  media=%s", image_b64[:20], media_type)

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
            body = resp.text[:400]
            log.error("Groq error %d: %s", resp.status_code, body)
            if attempt == 0:
                continue
            return err(f"Groq API returned error {resp.status_code}. Please try again.", 502)

        try:
            text   = resp.json()["choices"][0]["message"]["content"]
            parsed = extract_json(text)
        except Exception as exc:
            log.error("JSON parse error (attempt %d): %s", attempt + 1, exc)
            if attempt == 0:
                continue
            return err("Could not parse the AI response. Try a clearer photo.", 500)

        rows = normalize_rows(parsed.get("rows"))
        if not rows:
            if attempt == 0:
                log.warning("No rows returned; retrying.")
                continue
            return err("No data rows found in the photo. Ensure the image is clear and well-lit.", 422)

        log.info("Extracted %d rows, %d cols.", len(rows), max((len(r["values"]) for r in rows), default=0))
        return jsonify({"rows": rows})

    return err("Extraction failed after multiple attempts. Please use a clearer photo.", 500)


@app.route("/download-excel", methods=["POST"])
def download_excel():
    data   = request.get_json(silent=True) or {}
    output = build_excel(data.get("rows"))
    if output is None:
        return err("No rows provided.", 400)

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
    ready     = [d for d in documents if d.get("rows")]
    if not ready:
        return err("No completed documents provided.", 400)

    buf = BytesIO()
    used: set[str] = set()
    with zipfile.ZipFile(buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for i, doc in enumerate(ready, start=1):
            wb = build_excel(doc.get("rows"))
            if wb is None:
                continue
            base  = safe_filename(doc.get("filename"), f"factory-image-{i}")
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
        download_name="factory-excels.zip",
        mimetype="application/zip",
    )


# ── Entrypoint ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5050))
    log.info("Starting Factory Scanner (dev) on http://localhost:%d", port)
    app.run(host="0.0.0.0", debug=False, port=port)
