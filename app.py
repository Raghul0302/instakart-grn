import os
import re
import io
import json
import base64
import tempfile
import pdfplumber
import gspread
import sys

from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseUpload
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas as rl_canvas
from PyPDF2 import PdfReader, PdfWriter
from datetime import datetime

app = Flask(__name__)
CORS(app)

# =========================================================
# CONFIG
# =========================================================

SPREADSHEET_NAME  = "POD_OCR_DATA"
INPUT_SHEET        = "POD_INPUT"
OUTPUT_SHEET       = "POD_OUTPUT"
SUMMARY_SHEET      = "POD_SUMMARY"
DRIVE_FOLDER_ID    = "0AD4KdDfAgONUUk9PVA"
DRIVE_NOTIFY_EMAIL = "rraghul@ninjacart.com"

HEADERS = [
    "City", "Date", "Store",
    "Sheet Invoice ID", "PDF Invoice ID", "POD Status",
    "FSN", "Expected Qty", "Received Qty",
    "Damaged Qty", "Excess Qty", "Scanning Issue Qty", "Returned Qty",
    "Security Name", "HL Executive Name", "Inward Reg No",
    "Invoice Qty Box", "Received Qty Box", "Remark", "Drive Link", "Submitted At"
]

SUMMARY_HEADERS = [
    "City", "Date", "Store",
    "Sheet Invoice ID", "PDF Invoice ID", "POD Status",
    "Security Name", "HL Executive Name", "Inward Reg No",
    "Invoice Qty Box", "Received Qty Box", "Remark",
    "Drive Link", "Submitted At"
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
    try:
        summary = spreadsheet.worksheet(SUMMARY_SHEET)
    except Exception:
        summary = spreadsheet.add_worksheet(title=SUMMARY_SHEET, rows=5000, cols=20)
    return inp, out, summary, creds

# =========================================================
# GOOGLE DRIVE UPLOAD
# =========================================================

def upload_to_drive(creds, file_bytes: bytes, filename: str) -> str:
    """Upload PDF to Google Shared Drive folder, return link."""
    try:
        service = build("drive", "v3", credentials=creds)
        file_metadata = {
            "name"    : filename,
            "parents" : [DRIVE_FOLDER_ID],
            "driveId" : DRIVE_FOLDER_ID
        }
        media = MediaIoBaseUpload(io.BytesIO(file_bytes), mimetype="application/pdf")
        f = service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id, webViewLink",
            supportsAllDrives=True
        ).execute()
        # Share with office email instead of public
        service.permissions().create(
            fileId=f["id"],
            body={
                "type"        : "user",
                "role"        : "reader",
                "emailAddress": DRIVE_NOTIFY_EMAIL
            },
            supportsAllDrives=True,
            sendNotificationEmail=False
        ).execute()
        link = f.get("webViewLink", "")
        print(f"DRIVE UPLOAD SUCCESS: {link}", file=sys.stderr, flush=True)
        return link
    except Exception as e:
        print(f"DRIVE UPLOAD ERROR: {str(e)}", file=sys.stderr, flush=True)
        return ""

# =========================================================
# HELPERS
# =========================================================

def numeric_id(val):
    m = re.match(r'(\d+)', str(val).strip())
    return m.group(1) if m else str(val).strip()


def extract_invoice_id(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        full_text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    patterns = [
        r'Invoice[:\s#No\.]*\s*([A-Z0-9\-]{5,20})',
        r'Inv[:\s#No\.]*\s*([A-Z0-9\-]{5,20})',
        r'Invoice\s+No[:\s]*([A-Z0-9\-]{5,20})',
    ]
    for pat in patterns:
        m = re.search(pat, full_text, re.IGNORECASE)
        if m:
            val = m.group(1).strip()
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
    try:
        existing = output_sheet.get_all_values()
        if not existing:
            output_sheet.append_row(HEADERS)
        elif existing[0] != HEADERS:
            output_sheet.delete_rows(1)
            output_sheet.insert_row(HEADERS, 1)
    except Exception:
        pass


def ensure_summary_header(summary_sheet):
    try:
        existing = summary_sheet.get_all_values()
        if not existing:
            summary_sheet.append_row(SUMMARY_HEADERS)
        elif existing[0] != SUMMARY_HEADERS:
            summary_sheet.delete_rows(1)
            summary_sheet.insert_row(SUMMARY_HEADERS, 1)
    except Exception:
        pass


def is_already_submitted(output_sheet, invoice_id):
    try:
        all_vals = output_sheet.get_all_values()
        if len(all_vals) <= 1:
            return False
        num = numeric_id(invoice_id)
        for row in all_vals[1:]:
            if len(row) > 4:
                if numeric_id(str(row[4])) == num:
                    return True
        return False
    except Exception:
        return False


def b64_to_tempfile(b64_str):
    if not b64_str:
        return None
    try:
        raw = b64_str
        if "," in raw:
            raw = raw.split(",", 1)[1]
        img_bytes = base64.b64decode(raw)
        tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
        tmp.write(img_bytes)
        tmp.flush()
        tmp.close()
        return tmp.name
    except Exception:
        return None


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

    sec_img = b64_to_tempfile(data.get("security_sign", ""))
    hl_img  = b64_to_tempfile(data.get("hl_sign", ""))

    c.setFont("Helvetica-Bold", 13)
    c.setFillColor(BLACK)
    c.drawCentredString(W / 2, y, f"Inv No.: {data.get('invoice_id', '')}")
    y -= 13 * mm

    c.setFillColor(BLUE)
    c.rect(ML, y - 10 * mm, BW, 10 * mm, fill=1, stroke=0)
    c.setFillColor(WHITE)
    c.setFont("Helvetica-Bold", 10.5)
    c.drawCentredString(W / 2, y - 6.5 * mm, "INSTAKART SERVICES PVT. LTD.")
    y -= 10 * mm

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

    def draw_cell(cx, cy, cw, rh, label=None, value=None, img_path=None):
        c.setFillColor(WHITE)
        c.setStrokeColor(BLUE)
        c.setLineWidth(0.4)
        c.rect(cx, cy - rh, cw, rh, fill=1, stroke=1)
        if label:
            c.setFont("Helvetica-Bold", 7.5)
            c.setFillColor(BLUE)
            c.drawString(cx + 2.5 * mm, cy - 4.5 * mm, label)
        if value:
            c.setFont("Helvetica-Bold", 11)
            c.setFillColor(BLACK)
            c.drawString(cx + 2.5 * mm, cy - rh + 3 * mm, str(value))
        if img_path and os.path.exists(img_path):
            try:
                c.drawImage(
                    img_path,
                    cx + 1.5 * mm, cy - rh + 2 * mm,
                    cw - 3 * mm, rh - 6 * mm,
                    preserveAspectRatio=True, mask="auto"
                )
            except Exception:
                pass

    def draw_row(cells, rh_mm):
        nonlocal y
        rh = rh_mm * mm
        cx = ML
        for cell in cells:
            cw = BW * cell["w"]
            draw_cell(
                cx, y, cw, rh,
                label    = cell.get("label"),
                value    = cell.get("value"),
                img_path = cell.get("img_path"),
            )
            cx += cw
        y -= rh

    draw_row([
        {"label": "Date",               "value": data.get("date",""),     "w": 0.28},
        {"label": "Time",               "value": data.get("time",""),     "w": 0.22},
        {"label": "Invoice Qty / Box",  "value": data.get("inv_qty",""),  "w": 0.28},
        {"label": "Received Qty / Box", "value": data.get("rcvd_qty",""), "w": 0.22},
    ], 13)

    draw_row([
        {"label": "Inward Reg. Sl. No.", "value": data.get("inward_reg",""), "w": 1}
    ], 10)

    draw_row([
        {"label": "Security Name / ID", "value": data.get("security_name",""), "w": 0.38},
        {"label": "Security Sign.",     "img_path": sec_img,                   "w": 0.62},
    ], 22)

    draw_row([
        {"label": "HL Executive Name / ID", "value": data.get("hl_name",""), "w": 0.38},
        {"label": "HL Executive Sign.",     "img_path": hl_img,              "w": 0.62},
    ], 22)

    draw_row([
        {"label": "Remark", "value": data.get("remark",""), "w": 1}
    ], 11)

    c.setFillColor(BLUE)
    c.rect(ML, y - 9 * mm, BW, 9 * mm, fill=1, stroke=0)
    c.setFillColor(WHITE)
    c.setFont("Helvetica-Bold", 8)
    c.drawCentredString(W / 2, y - 6 * mm,
        "Received Physical Count & Quality subject to verification")

    c.setStrokeColor(BLUE)
    c.setLineWidth(0.3)
    c.rect(10, 10, W - 20, H - 20, fill=0, stroke=1)

    c.save()
    buf.seek(0)
    result = buf.read()

    for p in [sec_img, hl_img]:
        if p and os.path.exists(p):
            try: os.remove(p)
            except: pass

    return result


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
                "error": "Could not read Invoice ID from this PDF. Please check you uploaded the correct POD PDF."
            }), 400

        input_sheet, _, _, _creds = get_sheets()
        record = lookup_sheet_record(input_sheet, pdf_invoice_id)

        if not record:
            return jsonify({
                "valid"     : False,
                "invoice_id": pdf_invoice_id,
                "error"     : f"Invoice ID '{pdf_invoice_id}' not found in today's records. Please contact your manager."
            }), 404

        _, output_sheet, _, _creds = get_sheets()
        if is_already_submitted(output_sheet, pdf_invoice_id):
            return jsonify({
                "valid"     : False,
                "invoice_id": pdf_invoice_id,
                "error"     : f"⚠️ Invoice '{pdf_invoice_id}' has already been uploaded. Please contact your manager if this is a mistake."
            }), 409

        return jsonify({
            "valid"     : True,
            "invoice_id": pdf_invoice_id,
            "store"     : str(record.get("Store", "")),
            "city"      : str(record.get("City",  "")),
            "date"      : str(record.get("Date",  "")),
        })

    except Exception as e:
        return jsonify({"valid": False, "error": f"Error: {str(e)}"}), 500

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

        input_sheet, output_sheet, summary_sheet, creds = get_sheets()
        record = lookup_sheet_record(input_sheet, pdf_invoice_id)

        city             = seal_data.get("city",  record.get("City",  "") if record else "")
        date             = seal_data.get("date",  record.get("Date",  "") if record else "")
        store            = seal_data.get("store", record.get("Store", "") if record else "")
        sheet_invoice_id = record.get("Invoice_ID", pdf_invoice_id) if record else pdf_invoice_id
        submitted_at     = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        pod_status = "VALID POD" if numeric_id(sheet_invoice_id) == numeric_id(pdf_invoice_id) else "INVALID POD"

        # Build merged PDF
        seal_bytes   = build_seal_pdf({
            **seal_data,
            "invoice_id": pdf_invoice_id,
            "store"     : store,
            "city"      : city,
        })
        merged_bytes = merge_pdfs(pod_bytes, seal_bytes)
        fname = f"POD_{store.replace(' ','_')}_Inv{pdf_invoice_id}_{datetime.now().strftime('%d%m%Y')}.pdf"

        # Upload to Google Shared Drive
        drive_link = upload_to_drive(creds, merged_bytes, fname)

        # ── POD_OUTPUT sheet ──
        ensure_output_header(output_sheet)

        base_row = [city, date, store, sheet_invoice_id, pdf_invoice_id, pod_status]
        seal_row_no_link = [
            seal_data.get("security_name", ""),
            seal_data.get("hl_name", ""),
            seal_data.get("inward_reg", ""),
            seal_data.get("inv_qty", ""),
            seal_data.get("rcvd_qty", ""),
            seal_data.get("remark", ""),
        ]

        if not fsn_rows:
            output_sheet.append_row(
                base_row + ["NO RETURNS", 0, 0, 0, 0, 0, 0] +
                seal_row_no_link + [drive_link, submitted_at]
            )
        else:
            for i, r in enumerate(fsn_rows):
                row_drive_link = drive_link if i == 0 else ""
                output_sheet.append_row(
                    base_row + [
                        r["fsn"], r["expected_qty"], r["received_qty"],
                        r["damaged_qty"], r["excess_qty"], r["scanning_qty"], r["returned_qty"],
                    ] + seal_row_no_link + [row_drive_link, submitted_at]
                )

        # ── POD_SUMMARY sheet — one row per invoice ──
        ensure_summary_header(summary_sheet)
        summary_sheet.append_row([
            city, date, store,
            sheet_invoice_id, pdf_invoice_id, pod_status,
            seal_data.get("security_name", ""),
            seal_data.get("hl_name", ""),
            seal_data.get("inward_reg", ""),
            seal_data.get("inv_qty", ""),
            seal_data.get("rcvd_qty", ""),
            seal_data.get("remark", ""),
            drive_link,
            submitted_at,
        ])

        return send_file(
            io.BytesIO(merged_bytes),
            mimetype="application/pdf",
            as_attachment=True,
            download_name=fname
        )

    except Exception as e:
        print(f"SUBMIT ERROR: {str(e)}", file=sys.stderr, flush=True)
        return jsonify({"error": str(e)}), 500

    finally:
        if os.path.exists(tmp_path):
            os.remove(tmp_path)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port, debug=False)
