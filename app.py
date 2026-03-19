import os
import re
import json
import base64
import tempfile
from io import BytesIO

from flask import Flask, request, jsonify, send_file, send_from_directory
from flask_cors import CORS
from dotenv import load_dotenv
from openai import OpenAI
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

load_dotenv()

app = Flask(__name__, static_folder="public", static_url_path="")
CORS(app)

# Configure Groq via the OpenAI-compatible API
API_KEY = os.getenv("GROQ_API_KEY")
if not API_KEY:
    raise RuntimeError("GROQ_API_KEY not found in .env file")

client = OpenAI(
    api_key=API_KEY,
    base_url="https://api.groq.com/openai/v1",
)

# Default to Groq-hosted Llama 4 Scout, a vision-capable model with broader availability.
MODEL = os.getenv("GROQ_MODEL", "meta-llama/llama-4-scout-17b-16e-instruct")
VISION_MODELS = {
    "meta-llama/llama-4-scout-17b-16e-instruct",
    "meta-llama/llama-4-maverick-17b-128e-instruct",
}

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg", "webp", "bmp", "gif"}
MIME_MAP = {
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "webp": "image/webp",
    "bmp": "image/bmp",
    "gif": "image/gif",
}

# ─── Extraction Prompt ────────────────────────────────────────────────
# Each line is carefully crafted for financial report extraction.
# The prompt is structured in 4 sections:
#   1. ROLE      – sets the LLM's expertise context
#   2. TASK      – what exactly to extract
#   3. RULES     – how to handle financial report quirks
#   4. OUTPUT    – strict JSON format specification

EXTRACTION_PROMPT = """
You are a financial data extraction specialist.
Your task is to extract ALL tabular data from
the provided financial report screenshot.

This image is from a financial statement such as:
- Balance Sheet (Statement of Financial Position)
- Profit & Loss / Income Statement
- Cash Flow Statement
- Schedule of Notes to Accounts
- Statement of Changes in Equity

─── EXTRACTION RULES ───

1. EVERY row in the table must be captured — do NOT skip any line items,
   including section headers, subtotals, totals, and blank separator rows.

2. Preserve the EXACT hierarchy of the financial statement.
   Add a "Level" field to indicate depth:
   - "header"  → section headings (e.g. "Non-Current Assets", "Equity and Liabilities")
   - "item"    → individual line items (e.g. "Property, Plant and Equipment")
   - "subtotal"→ sub-totals   (e.g. "Total Non-Current Assets")
   - "total"   → grand totals (e.g. "Total Assets", "Total Equity and Liabilities")

3. For the "Note" or "Notes" column:
   - Keep the note reference number exactly as shown (e.g. "1", "2a", "III")
   - If a row has no note reference, use an empty string ""

4. NUMERIC VALUES — critical rules:
   - Remove any comma separators (e.g. "2,67,096" → "267096")
   - Keep negative numbers as negatives: "(1,234)" → "-1234"
   - Parenthesized numbers in financial reports are NEGATIVE values
   - Keep decimals if present (e.g. "12.50" stays "12.50")
   - If a cell shows "—" or "-" or is blank, use ""
   - Do NOT fabricate or estimate any numbers

5. COLUMN HEADERS:
   - Use the EXACT column headers as they appear in the document
   - For date-based columns, keep the full date (e.g. "As at 31st March, 2025")
   - Multi-line headers should be merged into one string

6. If there are MULTIPLE tables in the image (e.g. Assets on top, Equity & Liabilities below),
   combine them into a SINGLE flat array. Use the "Section" field to distinguish which
   part of the statement each row belongs to (e.g. "Assets", "Equity and Liabilities").

7. Currency symbols (₹, $, €, etc.) should NOT appear in numeric values.
   Mention the currency unit in a "Currency" field only if shown in the header.

─── OUTPUT FORMAT ───

Return ONLY a valid JSON array. No markdown, no explanation, no commentary.

Each object in the array must have EXACTLY these keys (plus any date columns):
{
  "Section": "which part of the financial statement",
  "Level": "header | item | subtotal | total",
  "Line Item": "the row label exactly as shown",
  "Notes": "note reference or empty string",
  "<Period Column 1>": "numeric value as string",
  "<Period Column 2>": "numeric value as string"
}

Where <Period Column 1>, <Period Column 2>, etc. are the actual date/period
column headers from the image (e.g. "As at 31st March, 2025").

CRITICAL: Output ONLY the JSON array. No text before or after it.
"""

def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def get_mime(filename):
    ext = filename.rsplit(".", 1)[1].lower()
    return MIME_MAP.get(ext, "image/png")


def extract_table_from_image(image_path, mime_type):
    """Send image to Groq and extract table data as JSON."""
    with open(image_path, "rb") as f:
        image_data = base64.b64encode(f.read()).decode("utf-8")

    if MODEL not in VISION_MODELS:
        raise RuntimeError(
            f"Configured model '{MODEL}' does not support image input on Groq. "
            "Set GROQ_MODEL to a vision-capable Groq model for screenshot extraction."
        )

    response = client.responses.create(
        model=MODEL,
        input=[
            {
                "role": "user",
                "content": [
                    {"type": "input_text", "text": EXTRACTION_PROMPT},
                    {
                        "type": "input_image",
                        "detail": "auto",
                        "image_url": f"data:{mime_type};base64,{image_data}",
                    },
                ],
            }
        ],
        temperature=1e-8,
        max_output_tokens=4096,
    )

    text = response.output_text
    print(f"--- RAW GROQ RESPONSE ---\n{text}\n--- END ---")

    return parse_json_from_text(text)


def parse_json_from_text(text):
    """Robustly extract a JSON array from model output that may contain extra text."""
    # Strip markdown code fences
    text = re.sub(r"```(?:json)?\s*", "", text).strip()

    # Try parsing with multiple strategies
    for candidate in _extract_candidates(text):
        try:
            data = json.loads(candidate)
            if isinstance(data, list):
                return data
            if isinstance(data, dict):
                return [data]
        except json.JSONDecodeError:
            pass

        # Try repairing common model JSON errors before giving up on this candidate
        repaired = _repair_json(candidate)
        try:
            data = json.loads(repaired)
            if isinstance(data, list):
                return data
            if isinstance(data, dict):
                return [data]
        except json.JSONDecodeError:
            pass

    raise json.JSONDecodeError("Could not find valid JSON in model response", text, 0)


def _extract_candidates(text):
    """Yield JSON candidate strings from text, from most to least specific."""
    # 1) Full text
    yield text

    # 2) Outermost JSON array [ ... ]
    match = re.search(r"\[[\s\S]*\]", text)
    if match:
        yield match.group()

    # 3) Outermost JSON object { ... }
    match = re.search(r"\{[\s\S]*\}", text)
    if match:
        yield match.group()


def _repair_json(text):
    """Fix common JSON errors produced by LLMs.

    Handles patterns like:
      "Total Assets": ""   →  used as a key, producing  "key1": "key2": "val"
    which is invalid JSON. We fix by collapsing double-colon key patterns.
    """
    # Fix  "key": "value": "value2"  →  "key": "value2"
    # This happens when the model puts a sub-header as a key-value pair
    text = re.sub(
        r'"([^"]*)":\s*"([^"]*)":\s*"([^"]*)"',
        r'"\1 \2": "\3"',
        text,
    )
    # Remove trailing commas before } or ]
    text = re.sub(r",\s*([}\]])", r"\1", text)
    return text


def create_excel(data):
    """Convert a list of dicts into a styled Excel workbook."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Extracted Data"

    if not data:
        return wb

    headers = list(data[0].keys())

    # --- Header style ---
    header_font = Font(name="Inter", bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill(start_color="4F46E5", end_color="4F46E5", fill_type="solid")
    header_align = Alignment(horizontal="center", vertical="center")
    thin_border = Border(
        left=Side(style="thin", color="E5E7EB"),
        right=Side(style="thin", color="E5E7EB"),
        top=Side(style="thin", color="E5E7EB"),
        bottom=Side(style="thin", color="E5E7EB"),
    )

    # Write headers
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border

    # --- Data style ---
    data_font = Font(name="Inter", size=10)
    data_align = Alignment(horizontal="left", vertical="center")
    alt_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

    for row_idx, row_data in enumerate(data, 2):
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=row_data.get(header, ""))
            cell.font = data_font
            cell.alignment = data_align
            cell.border = thin_border
            if row_idx % 2 == 0:
                cell.fill = alt_fill

    # Auto-fit column widths
    for col_idx, header in enumerate(headers, 1):
        max_len = len(str(header))
        for row_data in data:
            val = str(row_data.get(header, ""))
            max_len = max(max_len, len(val))
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = min(max_len + 4, 50)

    return wb


# ─── Routes ───────────────────────────────────────────────────────────

@app.route("/")
def index():
    return send_from_directory("public", "index.html")


@app.route("/preview", methods=["POST"])
def preview():
    """Upload image → return extracted JSON for in-page preview."""
    if "image" not in request.files:
        return jsonify({"error": "No image file provided"}), 400

    file = request.files["image"]
    if file.filename == "":
        return jsonify({"error": "No file selected"}), 400

    if not allowed_file(file.filename):
        return jsonify({"error": f"Unsupported file type. Use: {', '.join(ALLOWED_EXTENSIONS)}"}), 400

    # Save temp file
    tmp = tempfile.NamedTemporaryFile(delete=False, dir=UPLOAD_DIR, suffix=os.path.splitext(file.filename)[1])
    file.save(tmp.name)

    try:
        mime = get_mime(file.filename)
        data = extract_table_from_image(tmp.name, mime)
        return jsonify({"data": data})
    except json.JSONDecodeError:
        return jsonify({"error": "Could not parse table data from image. Try a clearer screenshot."}), 422
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        os.unlink(tmp.name)


@app.route("/convert", methods=["POST"])
def convert():
    """Upload image → return Excel file download."""
    if "image" not in request.files:
        return jsonify({"error": "No image file provided"}), 400

    file = request.files["image"]
    if file.filename == "":
        return jsonify({"error": "No file selected"}), 400

    if not allowed_file(file.filename):
        return jsonify({"error": f"Unsupported file type. Use: {', '.join(ALLOWED_EXTENSIONS)}"}), 400

    tmp = tempfile.NamedTemporaryFile(delete=False, dir=UPLOAD_DIR, suffix=os.path.splitext(file.filename)[1])
    file.save(tmp.name)

    try:
        mime = get_mime(file.filename)
        data = extract_table_from_image(tmp.name, mime)

        wb = create_excel(data)
        output = BytesIO()
        wb.save(output)
        output.seek(0)

        return send_file(
            output,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name="extracted_data.xlsx",
        )
    except json.JSONDecodeError:
        return jsonify({"error": "Could not parse table data from image. Try a clearer screenshot."}), 422
    except Exception as e:
        return jsonify({"error": str(e)}), 500
    finally:
        os.unlink(tmp.name)


if __name__ == "__main__":
    print("Server running at http://localhost:8000")
    app.run(debug=True, host="::", port=8000)
