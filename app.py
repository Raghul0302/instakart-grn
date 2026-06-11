import os
import re
import io
import json
import base64
import requests
import pdfplumber
import gspread

from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS
from google.oauth2.service_account import Credentials
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas as rl_canvas
from PyPDF2 import PdfReader, PdfWriter
from PIL import Image
from datetime import datetime

app = Flask(__name__)
CORS(app)

# =========================================================
# CONFIG
# =========================================================

SPREADSHEET_NAME = "POD_OCR_DATA"
INPUT_SHEET       = "POD_INPUT"
OUTPUT_SHEET      = "POD_OUTPUT"

HEADERS = [
    "City", "Date", "Store",
    "Sheet Invoice ID", "PDF Invoice ID", "POD Status",
    "FSN", "Expected Qty", "Received Qty",
    "Damaged Qty", "Excess Qty", "Scanning Issue Qty", "Returned Qty",
    "Security Name", "HL Executive Name", "Inward Reg No",
    "Invoice Qty Box", "Received Qty Box", "Remark", "Submitted At"
]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

os.makedirs("downloads", exist_ok=True)

# =========================================================
# GOOGLE SHEETS
# =========================================================

def get_sheets():
    import json as _json
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    if creds_json:
        creds_dict = _json.loads(creds_json)
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    else:
        creds = Credentials.from_service_account_file("credentials.json", scopes=SCOPES)
    gc          = gspread.authorize(creds)
    spreadsheet = gc.open(SPREADSHEET_NAME)
    inp         = spreadsheet.worksheet(INPUT_SHEET)
    try:
        out = spreadsheet.worksheet(OUTPUT_SHEET)
    except Exception:
        out = spreadsheet.add_worksheet(title=OUTPUT_SHEET, rows=5000, cols=25)
    return inp, out

# =========================================================
# HELPERS
# =========================================================

def numeric_id(val):
    m = re.match(r'(\d+)', str(val).strip())
    return m.group(1) if m else str(val).strip()


def extract_invoice_id(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    # Try multiple patterns
    patterns = [
        r'Invoice[:\s#No\.]*\s*([A-Z0-9\-]{5,20})',
        r'Inv[:\s#No\.]*\s*([A-Z0-9\-]{5,20})',
        r'Invoice\s+No[:\s]*([A-Z0-9\-]{5,20})',
    ]
    for pat in patterns:
        m = re.search(pat, full_text, re.IGNORECASE)
        if m:
            val = m.group(1).strip()
            # Must have at least 5 digits
            if re.search(r'\d{5}', val):
                return val
    return None


def parse_pod_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)

        def _find(pattern):
            m = re.search(pattern, full_text, re.IGNORECASE)
            return m.group(1).strip() if m else ""

        header = {
            "invoice_id"   : _find(r'Invoice[:\s#No\.]*\s*([A-Z0-9\-]{5,20})'),
            "invoice_date" : _find(r'Invoice Date:\s*(\S+)'),
            "warehouse_id" : _find(r'Warehouse ID:\s*(\S+)'),
            "status"       : _find(r'Status:\s*(\S+)'),
        }

        def _qty(cell):
            try: return int(str(cell).strip())
            except: return 0

        fsn_rows = []
        for page in pdf.pages:
            for table in (page.extract_tables() or []):
                for row in (table or []):
                    if not row or not row[0]: continue
                    fsn = str(row[0]).strip()
                    if not re.fullmatch(r'[A-Z0-9]{10,20}', fsn): continue
                    # Make sure row has enough columns
                    while len(row) < 8:
                        row.append("0")
                    fsn_rows.append({
                        "fsn"          : fsn,
                        "description"  : (row[1] or "").replace("\n", " ").strip(),
                        "expected_qty" : _qty(row[2]),
                        "received_qty" : _qty(row[3]),
                        "damaged_qty"  : _qty(row[4]),
                        "excess_qty"   : _qty(row[5]),
                        "scanning_qty" : _qty(row[6]),
                        "returned_qty" : _qty(row[7]) if len(row) > 7 else 0,
                    })

    return header, fsn_rows


def lookup_sheet_record(input_sheet, invoice_id):
    rows = input_sheet.get_all_records()
    num  = numeric_id(invoice_id)
    for row in rows:
        if numeric_id(str(row.get("Invoice_ID", ""))) == num:
            return row
    return None


def ensure_output_header(output_sheet):
    """Make sure the output sheet always has headers in row 1."""
    try:
        existing = output_sheet.get_all_values()
        if not existing or existing[0] != HEADERS:
            if not existing:
                output_sheet.append_row(HEADERS)
            else:
                output_sheet.insert_row(HEADERS, 1)
    except Exception:
        pass


# =========================================================
# SEAL PDF GENERATOR
# =========================================================

def build_seal_pdf(data: dict) -> bytes:
    buf = io.BytesIO()
    c   = rl_canvas.Canvas(buf, pagesize=A4)
    W, H = A4
    BLUE  = colors.HexColor("#1a3bb5")
    WHITE = colors.white
    BLACK = colors.HexColor("#141414")
    LBLUE = colors.HexColor("#eef1fb")

    ML = 25 * mm
    MR = 25 * mm
    BW = W - ML - MR
    y  = H - 25 * mm

    # Inv No
    c.setFont("Helvetica-Bold", 13)
    c.setFillColor(BLACK)
    c.drawCentredString(W / 2, y, f"Inv No.: {data.get('invoice_id', '—')}")
    y -= 13 * mm

    # Title bar
    c.setFillColor(BLUE)
    c.rect(ML, y - 10 * mm, BW, 10 * mm, fill=1, stroke=0)
    c.setFillColor(WHITE)
    c.setFont("Helvetica-Bold", 10.5)
    c.drawCentredString(W / 2, y - 6.5 * mm, "INSTAKART SERVICES PVT. LTD.")
    y -= 10 * mm

    # Sub header
    c.setFillColor(LBLUE)
    c.setStrokeColor(BLUE)
    c.setLineWidth(0.5)
    c.rect(ML, y - 8 * mm, BW, 8 * mm, fill=1, stroke=1)
    c.line(ML + BW / 2, y - 8 * mm, ML + BW / 2, y)
    c.setFillColor(BLUE)
    c.setFont("Helvetica-Bold", 9)
    c.drawString(ML + 3 * mm, y - 5 * mm, "GRN SEAL")
    c.drawRightString(ML + BW - 3 * mm, y - 5 * mm, data.get("store", ""))
    y -= 8 * mm

    def draw_row(cells, rh_mm):
        nonlocal y
        rh = rh_mm * mm
        cx = ML
        for cell in cells:
            cw = BW * cell["w"]
            c.setFillColor(WHITE)
            c.setStrokeColor(BLUE)
            c.setLineWidth(0.4)
            c.rect(cx, y - rh, cw, rh, fill=1, stroke=1)
            if cell.get("label"):
                c.setFont("Helvetica-Bold", 7.5)
                c.setFillColor(BLUE)
                c.drawString(cx + 2.5 * mm, y - 4.5 * mm, cell["label"])
            if cell.get("value"):
                c.setFont("Helvetica-Bold", 11)
                c.setFillColor(BLACK)
                c.drawString(cx + 2.5 * mm, y - rh + 3 * mm, str(cell["value"]))
            if cell.get("sign_b64"):
                try:
                    raw = cell["sign_b64"]
                    # strip data URI prefix if present
                    if "," in raw:
                        raw = raw.split(",", 1)[1]
                    img_data = base64.b64decode(raw)
                    img_buf  = io.BytesIO(img_data)
                    c.drawImage(
                        img_buf,
                        cx + 1.5 * mm, y - rh + 2 * mm,
                        cw - 3 * mm, rh - 6 * mm,
                        preserveAspectRatio=True, mask="auto"
                    )
                except Exception as ex:
                    # draw error text so we know it failed
                    c.setFont("Helvetica", 7)
                    c.setFillColor(colors.red)
                    c.drawString(cx + 2 * mm, y - rh + 3 * mm, f"[sig err: {ex}]")
            cx += cw
        y -= rh

    draw_row([
        {"label": "Date",               "value": data.get("date",""),     "w": 0.28},
        {"label": "Time",               "value": data.get("time",""),     "w": 0.22},
        {"label": "Invoice Qty / Box",  "value": data.get("inv_qty",""),  "w": 0.28},
        {"label": "Received Qty / Box", "value": data.get("rcvd_qty",""), "w": 0.22},
    ], 13)

    draw_row([{"label": "Inward Reg. Sl. No.", "value": data.get("inward_reg",""), "w": 1}], 10)

    draw_row([
        {"label": "Security Name / ID", "value": data.get("security_name",""),  "w": 0.38},
        {"label": "Security Sign.",     "sign_b64": data.get("security_sign",""), "w": 0.62},
    ], 22)

    draw_row([
        {"label": "HL Executive Name / ID", "value": data.get("hl_name",""),   "w": 0.38},
        {"label": "HL Executive Sign.",     "sign_b64": data.get("hl_sign",""), "w": 0.62},
    ], 22)

    draw_row([{"label": "Remark", "value": data.get("remark",""), "w": 1}], 11)

    # Footer
    c.setFillColor(BLUE)
    c.rect(ML, y - 9 * mm, BW, 9 * mm, fill=1, stroke=0)
    c.setFillColor(WHITE)
    c.setFont("Helvetica-Bold", 8)
    c.drawCentredString(W / 2, y - 6 * mm, "Received Physical Count & Quality subject to verification")
    y -= 9 * mm

    # Border
    c.setStrokeColor(BLUE)
    c.setLineWidth(0.3)
    c.rect(10, 10, W - 20, H - 20, fill=0, stroke=1)

    c.save()
    buf.seek(0)
    return buf.read()


def merge_pdfs(pod_bytes: bytes, seal_bytes: bytes) -> bytes:
    writer = PdfWriter()
    for reader in [PdfReader(io.BytesIO(pod_bytes)), PdfReader(io.BytesIO(seal_bytes))]:
        for page in reader.pages:
            writer.add_page(page)
    out_buf = io.BytesIO()
    writer.write(out_buf)
    out_buf.seek(0)
    return out_buf.read()


# =========================================================
# ROUTES
# =========================================================

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/validate", methods=["POST"])
def validate():
    if "pdf" not in request.files:
        return jsonify({"valid": False, "error": "No PDF uploaded."}), 400

    pdf_bytes = request.files["pdf"].read()
    tmp_path  = f"downloads/tmp_{datetime.now().strftime('%H%M%S%f')}.pdf"

    with open(tmp_path, "wb") as f:
        f.write(pdf_bytes)

    try:
        pdf_invoice_id = extract_invoice_id(tmp_path)

        if not pdf_invoice_id:
            return jsonify({
                "valid": False,
                "error": "Could not read Invoice ID from this PDF. Make sure you uploaded the correct POD PDF."
            }), 400

        input_sheet, _ = get_sheets()
        record = lookup_sheet_record(input_sheet, pdf_invoice_id)

        if not record:
            return jsonify({
                "valid"     : False,
                "invoice_id": pdf_invoice_id,
                "error"     : f"Invoice ID '{pdf_invoice_id}' is not in today's sheet. Please contact your manager."
            }), 404

        return jsonify({
            "valid"     : True,
            "invoice_id": pdf_invoice_id,
            "store"     : str(record.get("Store", "")),
            "city"      : str(record.get("City",  "")),
            "date"      : str(record.get("Date",  "")),
        })

    except Exception as e:
        return jsonify({"valid": False, "error": f"Server error: {str(e)}"}), 500

    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


@app.route("/submit", methods=["POST"])
def submit():
    seal_data_raw = request.form.get("seal_data", "{}")
    seal_data     = json.loads(seal_data_raw)

    if "pdf" not in request.files:
        return jsonify({"error": "No PDF uploaded."}), 400

    pod_bytes = request.files["pdf"].read()
    tmp_path  = f"downloads/sub_{datetime.now().strftime('%H%M%S%f')}.pdf"

    with open(tmp_path, "wb") as f:
        f.write(pod_bytes)

    try:
        header, fsn_rows = parse_pod_pdf(tmp_path)
        pdf_invoice_id   = header["invoice_id"] or seal_data.get("invoice_id", "UNKNOWN")

        input_sheet, output_sheet = get_sheets()
        record = lookup_sheet_record(input_sheet, pdf_invoice_id)

        city             = seal_data.get("city",  record.get("City",  "") if record else "")
        date             = seal_data.get("date",  record.get("Date",  "") if record else "")
        store            = seal_data.get("store", record.get("Store", "") if record else "")
        sheet_invoice_id = record.get("Invoice_ID", pdf_invoice_id) if record else pdf_invoice_id
        submitted_at     = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        pod_status = "VALID POD" if numeric_id(sheet_invoice_id) == numeric_id(pdf_invoice_id) else "INVALID POD"

        # Always ensure headers exist
        ensure_output_header(output_sheet)

        # Write FSN rows
        base_row = [
            city, date, store,
            sheet_invoice_id, pdf_invoice_id, pod_status,
        ]
        seal_row = [
            seal_data.get("security_name", ""),
            seal_data.get("hl_name", ""),
            seal_data.get("inward_reg", ""),
            seal_data.get("inv_qty", ""),
            seal_data.get("rcvd_qty", ""),
            seal_data.get("remark", ""),
            submitted_at,
        ]

        if not fsn_rows:
            output_sheet.append_row(base_row + ["NO RETURNS", 0, 0, 0, 0, 0, 0] + seal_row)
        else:
            for r in fsn_rows:
                output_sheet.append_row(base_row + [
                    r["fsn"], r["expected_qty"], r["received_qty"],
                    r["damaged_qty"], r["excess_qty"], r["scanning_qty"], r["returned_qty"],
                ] + seal_row)

        # Build and merge PDF
        seal_bytes   = build_seal_pdf({**seal_data, "invoice_id": pdf_invoice_id, "store": store, "city": city})
        merged_bytes = merge_pdfs(pod_bytes, seal_bytes)

        fname = f"GRN_{store.replace(' ','_')}_Inv{pdf_invoice_id}_{datetime.now().strftime('%d%m%Y')}.pdf"
        return send_file(io.BytesIO(merged_bytes), mimetype="application/pdf", as_attachment=True, download_name=fname)

    except Exception as e:
        return jsonify({"error": str(e)}), 500

    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=False)
