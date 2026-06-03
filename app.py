"""
Atlantic Packaging — Invoice Tools  v3.0
• Invoice Splitter   : Split batch PDFs into individual invoice PDFs
• Invoice Matcher    : Match invoices with POs, flatten & merge into single PDFs
• Invoice Codifier   : GL / Cost-Centre stamp on invoice PDFs
"""

import streamlit as st
import pdfplumber
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from io import BytesIO
import zipfile
import re
import copy
import json
import pandas as pd
from datetime import date, datetime
import openpyxl

# ── Optional OCR support ──────────────────────────────────────────────────────
try:
    import pytesseract
    from pdf2image import convert_from_bytes
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

# ── Optional pikepdf support (needed for Invoice Matcher) ─────────────────────
try:
    import pikepdf
    PIKEPDF_AVAILABLE = True
except ImportError:
    PIKEPDF_AVAILABLE = False

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Atlantic — Invoice Tools",
    layout="wide",
    page_icon="📄",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# DEFAULT DATA
# ─────────────────────────────────────────────────────────────────────────────
DEFAULT_PROVEEDORES = [
    {"prefijo": "ML", "vendor": "0101000430", "cc": "ML01"},
    {"prefijo": "EV", "vendor": "0101002430", "cc": "EV01"},
    {"prefijo": "MV", "vendor": "0101002220", "cc": "MV01"},
    {"prefijo": "MF", "vendor": "0101000430", "cc": "MF01"},
    {"prefijo": "MS", "vendor": "0101000430", "cc": "MS01"},
    {"prefijo": "MD", "vendor": "0101002430", "cc": "EV01"},
]

DEFAULT_GL_CODES = [
    {"codigo": "CM1023", "gl": "300052"}, {"codigo": "CM1026", "gl": "300051"},
    {"codigo": "CM1030", "gl": "300052"}, {"codigo": "CM1036", "gl": "300052"},
    {"codigo": "KGLB35", "gl": "300051"}, {"codigo": "LB2032", "gl": "300051"},
    {"codigo": "LB2035", "gl": "300051"}, {"codigo": "LB2042", "gl": "300051"},
    {"codigo": "LB2052", "gl": "300051"}, {"codigo": "PL2030", "gl": "300051"},
    {"codigo": "PRCM23", "gl": "300052"}, {"codigo": "WT56",   "gl": "300054"},
    {"codigo": "PRCM30", "gl": "300052"}, {"codigo": "CAWT33", "gl": "300054"},
    {"codigo": "KGLB42", "gl": "300051"}, {"codigo": "MDLU36", "gl": "300052"},
    {"codigo": "CAWT41", "gl": "300054"}, {"codigo": "CSLB42", "gl": "300051"},
    {"codigo": "003514", "gl": "300041"}, {"codigo": "003022", "gl": "300041"},
    {"codigo": "003502", "gl": "300041"}, {"codigo": "003024", "gl": "300041"},
    {"codigo": "DTCM23", "gl": "300052"}, {"codigo": "NDCM30", "gl": "300052"},
    {"codigo": "DTCM30", "gl": "300052"}, {"codigo": "LB2056", "gl": "300051"},
    {"codigo": "CM1033", "gl": "300052"}, {"codigo": "CAWT36", "gl": "300054"},
    {"codigo": "MDLU23", "gl": "300052"}, {"codigo": "CAWT25", "gl": "300054"},
    {"codigo": "003660", "gl": "300041"}, {"codigo": "003500", "gl": "300041"},
    {"codigo": "003771", "gl": "300041"}, {"codigo": "003021", "gl": "300041"},
    {"codigo": "001119", "gl": "300041"}, {"codigo": "002480", "gl": "300041"},
    {"codigo": "003675", "gl": "300041"}, {"codigo": "003501", "gl": "300041"},
    {"codigo": "003727", "gl": "300041"}, {"codigo": "003728", "gl": "300041"},
    {"codigo": "003729", "gl": "300041"}, {"codigo": "001777", "gl": "300041"},
    {"codigo": "002020", "gl": "300041"}, {"codigo": "002728", "gl": "300041"},
    {"codigo": "001912", "gl": "300041"}, {"codigo": "003366", "gl": "300041"},
    {"codigo": "003166", "gl": "300041"}, {"codigo": "002481", "gl": "300041"},
    {"codigo": "003691", "gl": "300041"}, {"codigo": "002488", "gl": "300041"},
    {"codigo": "003607", "gl": "300041"}, {"codigo": "002901", "gl": "300041"},
    {"codigo": "NDCM23", "gl": "300052"}, {"codigo": "WT36",   "gl": "300054"},
    {"codigo": "WT26",   "gl": "300054"}, {"codigo": "WT42",   "gl": "300054"},
    {"codigo": "WT31",   "gl": "300054"},
]

DEFAULT_USERS = ["ROC", "MLE", "PD"]
VENDOR_EXCEPCION = "0101000390"

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE INIT
# ─────────────────────────────────────────────────────────────────────────────
def init_state():
    defaults = {
        "proveedores":         copy.deepcopy(DEFAULT_PROVEEDORES),
        "gl_codes":            copy.deepcopy(DEFAULT_GL_CODES),
        "usuarios":            DEFAULT_USERS.copy(),
        "processed":           [],
        "stamp_x":             281,
        "stamp_y_top":         594,
        "stamp_w":             230,
        "stamp_h":             82,
        # Matcher state
        "matcher_results":     None,
        "matcher_zip":         None,
        "matcher_upload_key":  0,
        # Splitter state
        "splitter_results":    [],   # [{"filename", "pdf_bytes", "invoice_no", "cc", "bol", "pages", "warning"}]
        "splitter_zip":        None,
        "splitter_upload_key": 0,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

# ─────────────────────────────────────────────────────────────────────────────
# ── SPLITTER FUNCTIONS ────────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────────────────────

def _splitter_get_cc(raw_prefix: str) -> str:
    """
    Map the raw 2-char prefix from Customer Order No. to the CC used in filename.
    Looks up the vendor table: prefijo -> cc field -> take first 2 chars.
    e.g.  ML -> ML01 -> "ML"
          MD -> EV01 -> "EV"
    Falls back to the raw prefix if not found.
    """
    if not raw_prefix:
        return "??"
    key = raw_prefix.upper().strip()
    for row in st.session_state.proveedores:
        if str(row.get("prefijo", "")).upper().strip() == key:
            cc_val = str(row.get("cc", ""))
            return cc_val[:2].upper() if len(cc_val) >= 2 else cc_val.upper()
    return key  # fallback: use raw prefix as-is


def _splitter_extract_invoice_no(lines: list) -> str | None:
    """
    Invoice number appears on the line JUST BEFORE 'INVOICE No/No DE FACTURE'.
    Also handles it trailing on the same line.
    """
    for i, line in enumerate(lines):
        if "INVOICE NO" in line.upper() and "FACTURE" in line.upper():
            # Number at end of same line
            m = re.search(r"(\d{7,10})\s*$", line)
            if m:
                return m.group(1)
            # Number on any of the 4 preceding lines
            for j in range(max(0, i - 4), i):
                m = re.fullmatch(r"\d{6,10}", lines[j].strip())
                if m:
                    return lines[j].strip()
            break
    return None


def _splitter_extract_order_no(lines: list) -> str | None:
    """
    Customer Order No. appears in a line like:
      635108  100  ML11465  Dec 18, 2025  778713917 ...
    Returns the full order number e.g. 'ML11465'.
    """
    for line in lines:
        m = re.search(r"\d{6}\s+\d{2,3}\s+([A-Za-z]{2}\d{4,7})\b", line, re.IGNORECASE)
        if m:
            return m.group(1).upper()
    return None


def _splitter_extract_bol(lines: list) -> str | None:
    """
    Bill of Lading rules:
    • Two-line BOL — the second line contains the value to use:
        Layout A:  F0002775  1  CAWT25  ...  (line 1)
                   354192  195.331  MSF  4824  LB  ML11465  (line 2 ← USE THIS)
        Layout B:  F0002776  3  PRCM30  ...  (line 1)
                   24286651  606.06  MSF  18120  LB  ML11673  (line 2 ← USE THIS)
      Both layouts: line 2 starts with digits (5–9), then decimal/integer, then MSF.
      The only difference was digit length (6 vs 8) — now covered by \d{5,9}.
    • Single-line BOL:
        N0089112  8  LB2035  35# EnviroLiner ...  (alphanumeric ← USE THIS)
    Always use the LAST (second) BOL value when two exist.
    For multi-item invoices, use the BOL from the FIRST item only.
    """
    # Priority: second-line BOL — identifier followed by decimal+MSF
    # The identifier can be:
    #   • Pure digits   : 354192, 24286651  (layouts A & B)
    #   • Letter+digits : C249328           (layout C — letter-prefixed BOL)
    # [A-Z]? makes the leading letter optional, covering all three layouts.
    for line in lines:
        m = re.match(r"^([A-Za-z]{0,2}\d{5,9})\s+[\d,]+\.[\d]+\s+MSF\b", line.strip(), re.IGNORECASE)
        if m:
            return m.group(1).upper()

    # Fallback: alphanumeric single-line BOL (e.g. C0012857, N0089112)
    # Product code after qty can be alphabetic (CAWT25) OR numeric (001912) —
    # only require: letter + 5-9 digits, space, qty digits, space
    for line in lines:
        m = re.match(r"^([A-Za-z]\d{5,9})\s+\d+\s", line.strip(), re.IGNORECASE)
        if m:
            return m.group(1).upper()

    return None


def split_batch_pdf(pdf_bytes: bytes) -> list:
    """
    Split a multi-invoice Atlantic batch PDF into individual invoices.
    Returns list of dicts:
      { filename, pdf_bytes, invoice_no, cc, bol, pages (1-based), source_pages, warning }
    """
    reader   = PdfReader(BytesIO(pdf_bytes))
    invoices = []   # accumulated invoice dicts
    current  = None # {"number", "cc_raw", "bol", "pages": [0-based idx]}

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page_idx, page in enumerate(pdf.pages):
            text  = page.extract_text() or ""
            lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

            is_continuation = (
                "...Continued from previous page" in text
                or "Continued from previous page" in text
            )

            if is_continuation and current is not None:
                current["pages"].append(page_idx)
                continue

            inv_no = _splitter_extract_invoice_no(lines)

            if inv_no:
                if current is not None:
                    invoices.append(current)
                order_no = _splitter_extract_order_no(lines)
                cc_raw   = order_no[:2].upper() if order_no else None
                bol      = _splitter_extract_bol(lines)
                current  = {
                    "number":  inv_no,
                    "cc_raw":  cc_raw,
                    "bol":     bol,
                    "pages":   [page_idx],
                }
            elif current is not None:
                # Page without a clear invoice marker — attach to current
                current["pages"].append(page_idx)

    if current is not None:
        invoices.append(current)

    # Build output list
    results = []
    for inv in invoices:
        writer = PdfWriter()
        for p in inv["pages"]:
            writer.add_page(reader.pages[p])
        buf = BytesIO()
        writer.write(buf)

        inv_no = inv["number"]
        cc     = _splitter_get_cc(inv["cc_raw"]) if inv["cc_raw"] else "??"
        bol    = inv["bol"] or "??"
        warn   = None
        if inv["cc_raw"] is None:
            warn = "⚠️ Customer Order No. not detected — CC set to '??'"
        if inv["bol"] is None:
            warn = (warn or "") + "  ⚠️ Bill of Lading not detected — BOL set to '??'"

        results.append({
            "filename":     f"{inv_no} {cc} {bol}.pdf",
            "pdf_bytes":    buf.getvalue(),
            "invoice_no":   inv_no,
            "cc":           cc,
            "bol":          bol,
            "source_pages": [p + 1 for p in inv["pages"]],
            "page_count":   len(inv["pages"]),
            "warning":      warn,
        })

    return results


def make_splitter_zip(results: list) -> bytes:
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for item in results:
            zf.writestr(item["filename"], item["pdf_bytes"])
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────────────────────
# ── CODIFIER FUNCTIONS ────────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────────────────────

def get_vendor_cc(prefix: str):
    prefix = prefix.upper()
    for row in st.session_state.proveedores:
        if str(row["prefijo"]).upper() == prefix:
            return str(row["vendor"]), str(row["cc"])
    return None, None


def get_gl(product_code: str) -> str | None:
    if not product_code:
        return None
    code = product_code.upper().strip()
    for row in st.session_state.gl_codes:
        if str(row["codigo"]).upper().strip() == code:
            return str(row["gl"])
    return None


def extract_invoice_data(pdf_bytes: bytes, filename: str = "") -> dict:
    result = {
        "invoice_no": None, "customer_order": None, "cc_prefix": None,
        "product_code": None, "is_six": False, "raw_lines": [],
        "ocr_used": False, "error": None,
    }
    try:
        with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
            if not pdf.pages:
                result["error"] = "Empty PDF"
                return result
            text = pdf.pages[0].extract_text() or ""

        if len(text.strip()) < 50:
            if OCR_AVAILABLE:
                try:
                    images = convert_from_bytes(pdf_bytes, first_page=1, last_page=1, dpi=200)
                    text = pytesseract.image_to_string(images[0])
                    result["ocr_used"] = True
                except Exception as ocr_err:
                    result["error"] = f"OCR failed: {ocr_err}"
                    return result
            else:
                result["error"] = "Scanned PDF — OCR not available"
                return result

        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        result["raw_lines"] = lines
        text_upper = text.upper()

        for i, line in enumerate(lines):
            if "INVOICE NO" in line.upper() and "FACTURE" in line.upper():
                m = re.search(r"(\d{7,10})\s*$", line)
                if m:
                    result["invoice_no"] = m.group(1)
                    result["is_six"] = result["invoice_no"].startswith("6")
                    break
                for j in range(max(0, i - 5), i):
                    if re.fullmatch(r"\d{6,10}", lines[j]):
                        result["invoice_no"] = lines[j]
                        result["is_six"] = result["invoice_no"].startswith("6")
                        break
                break

        if not result["invoice_no"] and filename:
            m = re.match(r"(\d{6,10})", filename)
            if m:
                result["invoice_no"] = m.group(1)
                result["is_six"] = result["invoice_no"].startswith("6")

        for line in lines:
            m = re.search(r"\d{6}\s+\d{2,3}\s+\|?\s*([A-Za-z]{2}\d{4,7})\b", line, re.IGNORECASE)
            if m:
                result["customer_order"] = m.group(1).upper()
                result["cc_prefix"] = m.group(1)[:2].upper()
                break

        known_codes = sorted(
            [str(r["codigo"]).upper() for r in st.session_state.gl_codes],
            key=len, reverse=True,
        )
        for code in known_codes:
            if re.search(r"\b" + re.escape(code) + r"\b", text_upper):
                result["product_code"] = code
                break

    except Exception as e:
        result["error"] = str(e)

    return result


def create_stamp(user, vendor, cc, gl, coding_date, page_w, page_h, rotation=0):
    sw = st.session_state.stamp_w
    sh = st.session_state.stamp_h
    margin = 18
    date_str = (coding_date.strftime("%d/%m/%Y")
                if hasattr(coding_date, "strftime") else str(coding_date))
    stamp_lines = [
        f"POSTED BY: {user}",
        f"VENDOR: {vendor}",
        f"CC: {cc}  |  GL: {gl}",
        f"DATE: {date_str}",
    ]
    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=(page_w, page_h))
    if rotation in (90, 270):
        disp_w, disp_h = page_h, page_w
        disp_cx = disp_w / 2
        disp_cy = disp_h - margin - sh / 2
        if rotation == 90:
            cx_pdf, cy_pdf = page_w - disp_cy, disp_cx
            rot_angle = 90
        else:
            cx_pdf, cy_pdf = disp_cy, page_w - disp_cx
            rot_angle = 90
        c.saveState()
        c.translate(cx_pdf, cy_pdf)
        c.rotate(rot_angle)
        lx, ly = -sw / 2, -sh / 2
        c.setStrokeColorRGB(0.85, 0.0, 0.0)
        c.setFillColorRGB(1.0, 1.0, 1.0)
        c.setLineWidth(1.8)
        c.rect(lx, ly, sw, sh, fill=1)
        c.setFillColorRGB(0.85, 0.0, 0.0)
        c.setFont("Helvetica-Bold", 8.5)
        line_h = sh / 5.2
        tx, ty = lx + 10, ly + sh - line_h
        for i, line in enumerate(stamp_lines):
            c.drawString(tx, ty - i * line_h, line)
        c.restoreState()
    else:
        sx = st.session_state.stamp_x
        sy_top = st.session_state.stamp_y_top
        sy_bot = sy_top - sh
        c.setStrokeColorRGB(0.85, 0.0, 0.0)
        c.setFillColorRGB(1.0, 1.0, 1.0)
        c.setLineWidth(1.8)
        c.rect(sx, sy_bot, sw, sh, fill=1)
        c.setFillColorRGB(0.85, 0.0, 0.0)
        c.setFont("Helvetica-Bold", 8.5)
        line_h = sh / 5.2
        tx, ty = sx + 10, sy_bot + sh - line_h
        for i, line in enumerate(stamp_lines):
            c.drawString(tx, ty - i * line_h, line)
    c.save()
    packet.seek(0)
    return packet.read()


def stamp_pdf(original_bytes, stamp_bytes):
    reader = PdfReader(BytesIO(original_bytes))
    stamp_reader = PdfReader(BytesIO(stamp_bytes))
    stamp_page = stamp_reader.pages[0]
    writer = PdfWriter()
    first = reader.pages[0]
    first.merge_page(stamp_page)
    writer.add_page(first)
    for page in reader.pages[1:]:
        writer.add_page(page)
    out = BytesIO()
    writer.write(out)
    out.seek(0)
    return out.read()


def process_one(original_bytes, user, vendor, cc, gl, coding_date):
    reader = PdfReader(BytesIO(original_bytes))
    page = reader.pages[0]
    pw = float(page.mediabox.width)
    ph = float(page.mediabox.height)
    rotation = int(page.get("/Rotate", 0) or 0)
    stamp_bytes = create_stamp(user, vendor, cc, gl, coding_date, pw, ph, rotation)
    return stamp_pdf(original_bytes, stamp_bytes)


def make_zip(items):
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for item in items:
            zf.writestr(item["filename"], item["pdf_bytes"])
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────────────────────
# ── INVOICE MATCHER FUNCTIONS ─────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────────────────────

def extract_po_from_invoice_name(filename: str) -> str | None:
    """
    Invoice filename format: '{invoice_no} {cost_center} {PO_number}.pdf'
    e.g. '82196530 ML V0020978.pdf'  →  'V0020978'
    """
    name = re.sub(r'\.pdf$', '', filename, flags=re.IGNORECASE).strip()
    parts = name.split(' ')
    if len(parts) >= 3:
        return parts[2].strip()
    return None


def flatten_pdf(pdf_bytes: bytes) -> bytes:
    if PIKEPDF_AVAILABLE:
        try:
            inp = BytesIO(pdf_bytes)
            out = BytesIO()
            with pikepdf.open(inp, suppress_warnings=True) as pdf:
                pdf.save(out)
            out.seek(0)
            return out.read()
        except Exception:
            pass

    if OCR_AVAILABLE:
        try:
            from pdf2image import convert_from_bytes
            pages = convert_from_bytes(pdf_bytes, dpi=120)
            out = BytesIO()
            if len(pages) == 1:
                pages[0].save(out, format="PDF")
            else:
                pages[0].save(out, format="PDF", save_all=True,
                              append_images=pages[1:])
            out.seek(0)
            return out.read()
        except Exception:
            pass

    return pdf_bytes


def merge_two_pdfs(bytes1: bytes, bytes2: bytes) -> bytes:
    if PIKEPDF_AVAILABLE:
        out = BytesIO()
        with pikepdf.Pdf.new() as merged:
            with pikepdf.open(BytesIO(bytes1), suppress_warnings=True) as p1:
                merged.pages.extend(p1.pages)
            with pikepdf.open(BytesIO(bytes2), suppress_warnings=True) as p2:
                merged.pages.extend(p2.pages)
            merged.save(out)
        out.seek(0)
        return out.read()
    else:
        writer = PdfWriter()
        for b in (bytes1, bytes2):
            reader = PdfReader(BytesIO(b))
            for page in reader.pages:
                writer.add_page(page)
        out = BytesIO()
        writer.write(out)
        out.seek(0)
        return out.read()


def run_matching(invoice_files: list, po_files: list,
                 progress_callback=None) -> dict:
    po_lookup = {}
    for f in po_files:
        key = re.sub(r'\.pdf$', '', f.name, flags=re.IGNORECASE).strip().upper()
        po_lookup[key] = f.read()

    used_po_keys = set()
    matched = []
    pending = []
    total = len(invoice_files)

    for i, inv_file in enumerate(invoice_files):
        inv_bytes = inv_file.read()
        fname = inv_file.name

        if progress_callback:
            progress_callback(i, total, fname)

        po_id = extract_po_from_invoice_name(fname)

        if po_id is None:
            pending.append({
                "invoice_name": fname,
                "po_id":        "—",
                "reason":       "Invalid filename format (needs at least 3 space-separated parts)",
                "inv_bytes":    inv_bytes,
            })
            continue

        po_key = po_id.upper()
        if po_key in po_lookup:
            po_bytes = po_lookup[po_key]
            used_po_keys.add(po_key)
            flat_inv = flatten_pdf(inv_bytes)
            flat_po  = flatten_pdf(po_bytes)
            merged   = merge_two_pdfs(flat_inv, flat_po)
            matched.append({
                "invoice_name": fname,
                "po_name":      f"{po_id}.pdf",
                "po_id":        po_id,
                "merged_bytes": merged,
            })
        else:
            pending.append({
                "invoice_name": fname,
                "po_id":        po_id,
                "reason":       f"No PO file found for '{po_id}'",
                "inv_bytes":    inv_bytes,
            })

    unmatched_po = [
        name for name in po_lookup
        if name not in used_po_keys
    ]

    return {
        "matched":       matched,
        "pending":       pending,
        "unmatched_po":  unmatched_po,
    }


def make_matcher_zip(results: dict) -> bytes:
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for item in results["matched"]:
            zf.writestr(f"matched/{item['invoice_name']}", item["merged_bytes"])
        for item in results["pending"]:
            zf.writestr(f"pending/{item['invoice_name']}", item["inv_bytes"])
    buf.seek(0)
    return buf.read()


# ─────────────────────────────────────────────────────────────────────────────
# CSS STYLES
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
.stTabs [data-baseweb="tab"] { font-size: 15px; font-weight: 600; padding: 8px 18px; }
.inv-card {
    background: #f8f9fa; border-radius: 10px; padding: 14px 18px;
    margin-bottom: 8px; border-left: 5px solid #ccc;
}
.inv-ok   { border-left-color: #28a745 !important; }
.inv-warn { border-left-color: #ffc107 !important; }
.inv-err  { border-left-color: #dc3545 !important; }
.stamp-preview {
    border: 2px solid #cc0000; padding: 8px 12px;
    display: inline-block; background: white;
    font-family: monospace; font-weight: bold; color: #cc0000;
    font-size: 13px; line-height: 1.6; border-radius: 3px;
}
.match-row {
    background: #f0fff4; border-radius: 8px; padding: 10px 14px;
    margin-bottom: 6px; border-left: 4px solid #28a745;
    font-size: 14px;
}
.pending-row {
    background: #fff8f0; border-radius: 8px; padding: 10px 14px;
    margin-bottom: 6px; border-left: 4px solid #ffc107;
    font-size: 14px;
}
.stat-box {
    background: white; border-radius: 10px; padding: 16px 20px;
    text-align: center; box-shadow: 0 1px 4px rgba(0,0,0,0.08);
    border: 1px solid #e9ecef;
}
.stat-num  { font-size: 36px; font-weight: 700; line-height: 1.1; }
.stat-lbl  { font-size: 13px; color: #6c757d; margin-top: 2px; }
.green { color: #28a745; }
.amber { color: #e08000; }
.red   { color: #dc3545; }
.blue  { color: #0d6efd; }
.split-row {
    background: #f8f9fa; border-radius: 8px; padding: 10px 14px;
    margin-bottom: 6px; border-left: 4px solid #0d6efd;
    font-size: 14px;
}
.split-warn {
    border-left-color: #ffc107 !important;
    background: #fffdf0 !important;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Work Session")

    user_list = st.session_state.usuarios + ["✏️ Other..."]
    sel_idx = st.selectbox(
        "👤 Posted By",
        range(len(user_list)),
        format_func=lambda x: user_list[x],
    )
    if user_list[sel_idx] == "✏️ Other...":
        current_user = st.text_input("Name:", placeholder="Enter name")
    else:
        current_user = user_list[sel_idx]

    coding_date = st.date_input("📅 Coding Date", value=date.today())

    st.divider()

    n_proc = len(st.session_state.processed)
    st.metric("Coded Invoices", n_proc)

    if n_proc > 0:
        zip_bytes = make_zip(st.session_state.processed)
        st.download_button(
            "⬇️ Download ZIP (all)",
            data=zip_bytes,
            file_name=f"invoices_{date.today().strftime('%Y%m%d')}.zip",
            mime="application/zip",
            use_container_width=True,
        )
        if st.button("🗑️ Clear all results", use_container_width=True, type="secondary"):
            st.session_state.processed = []
            st.rerun()

    st.divider()

    with st.expander("💾 Save / Load Database"):
        db_export = {
            "proveedores": st.session_state.proveedores,
            "gl_codes":    st.session_state.gl_codes,
            "usuarios":    st.session_state.usuarios,
        }
        st.download_button(
            "⬇️ Export DB (JSON)",
            data=json.dumps(db_export, indent=2, ensure_ascii=False),
            file_name="invoice_db.json",
            mime="application/json",
            use_container_width=True,
        )
        db_upload = st.file_uploader("📥 Import DB (JSON)", type=["json"], key="db_import")
        if db_upload:
            try:
                db = json.loads(db_upload.read())
                if "proveedores" in db: st.session_state.proveedores = db["proveedores"]
                if "gl_codes"    in db: st.session_state.gl_codes    = db["gl_codes"]
                if "usuarios"    in db: st.session_state.usuarios     = db["usuarios"]
                st.success("✅ Database loaded")
            except Exception as e:
                st.error(f"Error: {e}")

# ─────────────────────────────────────────────────────────────────────────────
# MAIN HEADER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<h1 style='margin-bottom:0'>📄 Atlantic — Invoice Tools</h1>
<p style='color:gray;margin-top:4px'>Invoice Splitter &nbsp;·&nbsp; Invoice &amp; PO Matcher &nbsp;·&nbsp; Invoice Codifier</p>
""", unsafe_allow_html=True)

tab_split, tab_match, tab_proc, tab_res, tab_db, tab_cfg = st.tabs([
    "✂️  Invoice Splitter",
    "🔗  Invoice Matcher",
    "📤  Process Invoices",
    "📋  Results",
    "🗄️  Database",
    "⚙️  Settings",
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — INVOICE SPLITTER
# ══════════════════════════════════════════════════════════════════════════════
with tab_split:
    st.subheader("✂️ Split Batch Invoice PDFs")
    st.markdown(
        "Upload one or more **batch PDFs** from Atlantic (each may contain multiple invoices). "
        "The tool will detect each invoice automatically, split them into individual PDFs, "
        "and name each file as **`{Invoice No} {CC} {Bill of Lading}.pdf`** — "
        "ready to use directly in the Matcher tab."
    )

    col_up, col_clear = st.columns([5, 1])
    with col_up:
        split_uploads = st.file_uploader(
            "Drag or select one or more batch PDFs",
            type=["pdf"],
            accept_multiple_files=True,
            key=f"split_uploader_{st.session_state.splitter_upload_key}",
        )
    with col_clear:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🗑️ Clear", use_container_width=True, key="split_clear"):
            st.session_state.splitter_results  = []
            st.session_state.splitter_zip      = None
            st.session_state.splitter_upload_key += 1
            st.rerun()

    if split_uploads:
        do_split = st.button("✂️ Split Invoices", type="primary", use_container_width=False)

        if do_split:
            st.session_state.splitter_results = []
            st.session_state.splitter_zip     = None
            all_results = []
            prog = st.progress(0, text="Processing…")

            for f_idx, f in enumerate(split_uploads):
                prog.progress((f_idx) / len(split_uploads),
                              text=f"Splitting {f.name}…")
                try:
                    batch_results = split_batch_pdf(f.read())
                    for r in batch_results:
                        r["source_file"] = f.name
                    all_results.extend(batch_results)
                except Exception as e:
                    st.error(f"Error processing **{f.name}**: {e}")

            prog.progress(1.0, text="✅ Done")
            st.session_state.splitter_results = all_results
            if all_results:
                st.session_state.splitter_zip = make_splitter_zip(all_results)
            st.rerun()

    # ── Results display ───────────────────────────────────────────────────────
    results = st.session_state.splitter_results

    if results:
        n_ok   = sum(1 for r in results if not r["warning"])
        n_warn = sum(1 for r in results if r["warning"])

        # Summary stats
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(
                f"<div class='stat-box'>"
                f"<div class='stat-num blue'>{len(results)}</div>"
                f"<div class='stat-lbl'>Invoices detected</div></div>",
                unsafe_allow_html=True,
            )
        with c2:
            st.markdown(
                f"<div class='stat-box'>"
                f"<div class='stat-num green'>{n_ok}</div>"
                f"<div class='stat-lbl'>Ready to download</div></div>",
                unsafe_allow_html=True,
            )
        with c3:
            st.markdown(
                f"<div class='stat-box'>"
                f"<div class='stat-num amber'>{n_warn}</div>"
                f"<div class='stat-lbl'>Need review</div></div>",
                unsafe_allow_html=True,
            )

        st.markdown("<br>", unsafe_allow_html=True)

        # Download all ZIP
        if st.session_state.splitter_zip:
            st.download_button(
                f"⬇️ Download all as ZIP ({len(results)} invoices)",
                data=st.session_state.splitter_zip,
                file_name=f"split_invoices_{date.today().strftime('%Y%m%d')}.zip",
                mime="application/zip",
                type="primary",
            )

        st.divider()

        # Per-invoice rows
        for item in results:
            row_cls = "split-row split-warn" if item["warning"] else "split-row"
            icon    = "⚠️" if item["warning"] else "✅"

            col_info, col_dl = st.columns([5, 1])
            with col_info:
                pages_str = ", ".join(str(p) for p in item["source_pages"])
                page_lbl  = f"{item['page_count']} page{'s' if item['page_count'] > 1 else ''}"
                st.markdown(
                    f"<div class='{row_cls}'>"
                    f"{icon} <b>{item['filename']}</b>"
                    f"&nbsp;&nbsp;|&nbsp;&nbsp;"
                    f"Invoice: <code>{item['invoice_no']}</code> &nbsp; "
                    f"CC: <code>{item['cc']}</code> &nbsp; "
                    f"BOL: <code>{item['bol']}</code> &nbsp; "
                    f"· {page_lbl} (source p. {pages_str})"
                    f"{'<br><span style=\"color:#e08000;font-size:12px\">' + item['warning'] + '</span>' if item['warning'] else ''}"
                    f"</div>",
                    unsafe_allow_html=True,
                )
            with col_dl:
                st.download_button(
                    "⬇️",
                    data=item["pdf_bytes"],
                    file_name=item["filename"],
                    mime="application/pdf",
                    key=f"spdl_{item['filename']}_{item['invoice_no']}",
                    use_container_width=True,
                    help=f"Download {item['filename']}",
                )

    elif not split_uploads:
        st.info("📂 Upload one or more Atlantic batch PDFs to get started.")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — INVOICE MATCHER
# ══════════════════════════════════════════════════════════════════════════════
with tab_match:
    st.subheader("🔗 Invoice & PO Matcher")
    st.markdown(
        "Upload your invoice PDFs and your PO (Purchase Order) PDFs. "
        "The tool will match them by PO number, flatten both documents to ensure "
        "consistent rendering, merge them into a single PDF, and package everything for download."
    )

    if not PIKEPDF_AVAILABLE:
        st.warning(
            "⚠️ **pikepdf** is not installed. Flattening will use rasterization (pdf2image) as fallback. "
            "Add `pikepdf` to `requirements.txt` for best results."
        )

    st.info(
        "📋 **Expected invoice filename format:** `{Invoice No} {Cost Centre} {PO Number}.pdf`  \n"
        "Example: `82196530 ML V0020978.pdf`  →  will search for PO file `V0020978.pdf`"
    )

    col_inv, col_po = st.columns(2)
    with col_inv:
        st.markdown("**📄 Invoice PDFs**")
        inv_files = st.file_uploader(
            "Upload invoices",
            type=["pdf"],
            accept_multiple_files=True,
            key=f"match_inv_{st.session_state.matcher_upload_key}",
            label_visibility="collapsed",
        )
        if inv_files:
            st.caption(f"{len(inv_files)} invoice(s) loaded")

    with col_po:
        st.markdown("**📦 PO PDFs**")
        po_files = st.file_uploader(
            "Upload POs",
            type=["pdf"],
            accept_multiple_files=True,
            key=f"match_po_{st.session_state.matcher_upload_key}",
            label_visibility="collapsed",
        )
        if po_files:
            st.caption(f"{len(po_files)} PO(s) loaded")

    col_run, col_clr, _ = st.columns([2, 2, 5])
    with col_run:
        run_match = st.button(
            "🔗 Run Matching",
            type="primary",
            use_container_width=True,
            disabled=not bool(inv_files),
        )
    with col_clr:
        if st.button("🗑️ Clear results", use_container_width=True):
            st.session_state.matcher_results    = None
            st.session_state.matcher_zip        = None
            st.session_state.matcher_upload_key += 1
            st.rerun()

    if run_match and inv_files:
        if not po_files:
            st.warning("⚠️ No PO files uploaded — all invoices will be marked as pending.")

        prog_bar  = st.progress(0, text="Matching…")
        prog_text = st.empty()

        def _progress(cur, tot, fname):
            prog_bar.progress((cur + 1) / tot, text=f"Processing {fname}…")
            prog_text.caption(f"({cur + 1}/{tot})")

        results = run_matching(inv_files, po_files or [], progress_callback=_progress)
        prog_bar.progress(1.0, text="✅ Done")
        prog_text.empty()

        st.session_state.matcher_results = results
        st.session_state.matcher_zip     = make_matcher_zip(results)
        st.rerun()

    # ── Results ────────────────────────────────────────────────────────────────
    if st.session_state.matcher_results:
        results     = st.session_state.matcher_results
        matched     = results["matched"]
        pending     = results["pending"]
        unmatched_po = results["unmatched_po"]

        # Stats
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown(
                f"<div class='stat-box'><div class='stat-num green'>{len(matched)}</div>"
                f"<div class='stat-lbl'>Matched</div></div>",
                unsafe_allow_html=True,
            )
        with c2:
            st.markdown(
                f"<div class='stat-box'><div class='stat-num amber'>{len(pending)}</div>"
                f"<div class='stat-lbl'>Unmatched invoices</div></div>",
                unsafe_allow_html=True,
            )
        with c3:
            st.markdown(
                f"<div class='stat-box'><div class='stat-num blue'>{len(unmatched_po)}</div>"
                f"<div class='stat-lbl'>Unused POs</div></div>",
                unsafe_allow_html=True,
            )

        st.markdown("<br>", unsafe_allow_html=True)

        if st.session_state.matcher_zip:
            st.download_button(
                "⬇️ Download all results (ZIP)",
                data=st.session_state.matcher_zip,
                file_name=f"matched_invoices_{date.today().strftime('%Y%m%d')}.zip",
                mime="application/zip",
                type="primary",
            )

        st.divider()

        if matched:
            with st.expander(f"✅ Matched ({len(matched)})", expanded=True):
                for item in matched:
                    col1, col2 = st.columns([5, 1])
                    with col1:
                        st.markdown(
                            f"<div class='match-row'>"
                            f"✅ <b>{item['invoice_name']}</b>"
                            f"&nbsp;&nbsp;+&nbsp;&nbsp;"
                            f"📦 <code>{item['po_name']}</code>"
                            f"</div>",
                            unsafe_allow_html=True,
                        )
                    with col2:
                        st.download_button(
                            "⬇️",
                            data=item["merged_bytes"],
                            file_name=item["invoice_name"],
                            mime="application/pdf",
                            key=f"mdl_{item['invoice_name']}",
                            use_container_width=True,
                        )

        if pending:
            with st.expander(f"⚠️ Unmatched Invoices ({len(pending)})", expanded=True):
                st.caption("These invoices had no corresponding PO file.")
                for item in pending:
                    st.markdown(
                        f"<div class='pending-row'>"
                        f"📄 <b>{item['invoice_name']}</b>"
                        f"&nbsp;&nbsp;|&nbsp;&nbsp;"
                        f"PO searched: <code>{item['po_id']}</code>"
                        f"&nbsp;&nbsp;—&nbsp;&nbsp;"
                        f"{item['reason']}"
                        f"</div>",
                        unsafe_allow_html=True,
                    )

        if unmatched_po:
            with st.expander(f"📦 Unused POs ({len(unmatched_po)})"):
                st.caption("These PO files were uploaded but no invoice referenced them.")
                for po_name in unmatched_po:
                    st.markdown(f"- `{po_name}.pdf`")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — PROCESS INVOICES (CODIFIER)
# ══════════════════════════════════════════════════════════════════════════════
with tab_proc:
    st.subheader("Upload Invoice PDFs")

    if "upload_key" not in st.session_state:
        st.session_state.upload_key = 0

    col_up, col_clear = st.columns([5, 1])
    with col_up:
        uploaded = st.file_uploader(
            "Drag or select one or more PDFs",
            type=["pdf"],
            accept_multiple_files=True,
            help="Supports bulk upload. Only the first page of each invoice is stamped.",
            key=f"uploader_{st.session_state.upload_key}",
        )
    with col_clear:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🗑️ Clear\nupload", use_container_width=True,
                     help="Remove all loaded files to upload a new batch"):
            st.session_state.upload_key += 1
            st.rerun()

    if not uploaded:
        st.info("📂 Upload invoices to get started. You can select multiple files at once.")
    else:
        st.success(f"✅ **{len(uploaded)} file(s)** loaded — analyzing…")
        st.divider()

        invoices_ui = []
        for idx, f in enumerate(uploaded):
            raw = f.read()
            data = extract_invoice_data(raw, f.name)
            data["filename"] = f.name
            data["raw_bytes"] = raw

            inv_no = data.get("invoice_no") or ""
            is_six = data.get("is_six", False)

            if is_six:
                data["vendor_auto"] = VENDOR_EXCEPCION
                prefix = data.get("cc_prefix")
                _, c = get_vendor_cc(prefix) if prefix else (None, None)
                data["cc_auto"]  = c
                data["needs_cc"] = (c is None)
            else:
                prefix = data.get("cc_prefix")
                v, c = get_vendor_cc(prefix) if prefix else (None, None)
                data["vendor_auto"] = v
                data["cc_auto"]     = c
                data["needs_cc"]    = (v is None or c is None)

            data["gl_auto"] = get_gl(data.get("product_code"))
            invoices_ui.append(data)

        st.subheader("Review — Extracted Data")

        resolved_cc     = {}
        resolved_vendor = {}
        resolved_gl     = {}

        for idx, inv in enumerate(invoices_ui):
            all_ok = (
                inv.get("invoice_no")
                and not inv["needs_cc"]
                and inv.get("gl_auto")
            )
            has_err = inv.get("error") or (
                not inv.get("invoice_no") and not inv.get("customer_order")
            )
            needs_input = inv["needs_cc"] or not inv.get("gl_auto")

            if has_err:
                icon, card_cls = "❌", "inv-card inv-err"
                expanded = True
            elif needs_input:
                icon, card_cls = "⚠️", "inv-card inv-warn"
                expanded = True
            else:
                icon, card_cls = "✅", "inv-card inv-ok"
                expanded = False

            with st.expander(f"{icon}  {inv['filename']}", expanded=expanded):
                c1, c2, c3 = st.columns([1.2, 1.2, 1])

                with c1:
                    st.markdown("**📑 Extracted from PDF**")
                    st.write(f"Invoice No: `{inv.get('invoice_no') or '—'}`")
                    st.write(f"Customer Order: `{inv.get('customer_order') or '—'}`")
                    st.write(f"CC Prefix: `{inv.get('cc_prefix') or '—'}`")
                    st.write(f"Product Code: `{inv.get('product_code') or '—'}`")
                    if inv.get("is_six"):
                        st.info("⚡ Type-6 invoice — vendor exception")
                    if inv.get("error"):
                        st.error(f"Error: {inv['error']}")
                    if inv.get("ocr_used"):
                        st.info("🔍 Text extracted via OCR (scanned PDF)")

                with c2:
                    st.markdown("**🏷️ Vendor / Cost Centre**")
                    if inv["is_six"]:
                        st.write(f"Vendor (fixed): `{VENDOR_EXCEPCION}`")
                        resolved_vendor[idx] = VENDOR_EXCEPCION
                        if inv["cc_auto"]:
                            st.success(f"CC: `{inv['cc_auto']}`")
                            resolved_cc[idx] = inv["cc_auto"]
                        else:
                            st.warning("CC prefix not detected — select manually")
                            cc_opts = sorted(set(r["cc"] for r in st.session_state.proveedores))
                            sel_cc = st.selectbox("Select CC:", cc_opts, key=f"cc6_{idx}")
                            resolved_cc[idx] = sel_cc
                    elif inv["vendor_auto"] and inv["cc_auto"]:
                        st.success(f"Vendor: `{inv['vendor_auto']}`")
                        st.success(f"CC: `{inv['cc_auto']}`")
                        resolved_vendor[idx] = inv["vendor_auto"]
                        resolved_cc[idx]     = inv["cc_auto"]
                    else:
                        st.warning(f"Prefix `{inv.get('cc_prefix')}` not found")
                        all_opts = sorted(set(r["cc"] for r in st.session_state.proveedores))
                        sel_cc = st.selectbox("Manual CC:", all_opts, key=f"ccman_{idx}")
                        sel_vendor = next(
                            (r["vendor"] for r in st.session_state.proveedores if r["cc"] == sel_cc),
                            "UNKNOWN"
                        )
                        resolved_cc[idx]     = sel_cc
                        resolved_vendor[idx] = sel_vendor

                with c3:
                    st.markdown("**📊 GL Account**")
                    if inv.get("gl_auto"):
                        st.success(f"GL: `{inv['gl_auto']}`")
                        resolved_gl[idx] = inv["gl_auto"]
                    else:
                        st.warning("GL not detected automatically")
                        gl_opts = sorted(set(r["gl"] for r in st.session_state.gl_codes))
                        sel_gl = st.selectbox("Manual GL:", gl_opts, key=f"glman_{idx}")
                        resolved_gl[idx] = sel_gl

                cc_prev  = resolved_cc.get(idx, "??")
                gl_prev  = resolved_gl.get(idx, "??")
                vd_prev  = resolved_vendor.get(idx, "??")
                usr_prev = current_user or "???"
                date_prev = coding_date.strftime("%d/%m/%Y")
                st.markdown(f"""
                <div style='margin-top:10px'>
                <p style='margin-bottom:4px; color:gray; font-size:12px'>👁️ Stamp preview:</p>
                <div class='stamp-preview'>
                POSTED BY: {usr_prev}<br>
                VENDOR: {vd_prev}<br>
                CC: {cc_prev}&nbsp;&nbsp;|&nbsp;&nbsp;GL: {gl_prev}<br>
                DATE: {date_prev}
                </div></div>
                """, unsafe_allow_html=True)

        st.divider()
        col_btn, col_info = st.columns([1, 3])
        with col_btn:
            do_process = st.button(
                "🚀 Process Invoices",
                type="primary",
                use_container_width=True,
                disabled=not bool(current_user),
            )
        with col_info:
            if not current_user:
                st.warning("⚠️ Select a responsible user (Posted By) in the sidebar before processing.")

        if do_process and current_user:
            progress = st.progress(0, text="Starting…")
            errors = []
            for idx, inv in enumerate(invoices_ui):
                fname = inv["filename"]
                progress.progress((idx + 1) / len(invoices_ui), text=f"Processing {fname}…")
                cc     = resolved_cc.get(idx, "???")
                vendor = resolved_vendor.get(idx, "???")
                gl     = resolved_gl.get(idx, "???")
                try:
                    stamped = process_one(inv["raw_bytes"], current_user, vendor, cc, gl, coding_date)
                    st.session_state.processed.append({
                        "filename":       fname,
                        "original_bytes": inv["raw_bytes"],
                        "pdf_bytes":      stamped,
                        "invoice_no":     inv.get("invoice_no"),
                        "vendor":         vendor,
                        "cc":             cc,
                        "gl":             gl,
                        "user":           current_user,
                        "date":           coding_date.strftime("%d/%m/%Y"),
                        "date_obj":       coding_date,
                        "ts":             datetime.now().strftime("%Y-%m-%d %H:%M"),
                    })
                except Exception as e:
                    errors.append(f"{fname}: {e}")
            progress.progress(1.0, text="✅ Done")
            if errors:
                for err in errors:
                    st.error(err)
            else:
                st.success(f"🎉 **{len(invoices_ui)} invoice(s)** coded successfully. Go to **Results** to download them.")
                st.balloons()


# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — RESULTS
# ══════════════════════════════════════════════════════════════════════════════
with tab_res:
    if not st.session_state.processed:
        st.info("📭 No invoices processed yet. Go to **Process Invoices** to begin.")
    else:
        n = len(st.session_state.processed)
        st.subheader(f"Processed Invoices — {n} file(s)")

        col_dl, col_del, _ = st.columns([2, 2, 5])
        with col_dl:
            zip_all = make_zip(st.session_state.processed)
            st.download_button(
                f"⬇️ Download ZIP ({n} invoices)",
                data=zip_all,
                file_name=f"coded_invoices_{date.today().strftime('%Y%m%d')}.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True,
            )
        with col_del:
            if st.button("🗑️ Delete all", use_container_width=True):
                st.session_state.processed = []
                st.rerun()

        st.divider()

        with st.expander("📅 Modify Coding Date"):
            st.caption("Use this if the invoice could not be registered on the same day it was coded.")
            col_nd, col_sel = st.columns([1, 2])
            with col_nd:
                new_date = st.date_input("New date:", value=date.today(), key="new_date_picker")
            with col_sel:
                inv_labels = [f"{item['filename']}  ({item['date']})" for item in st.session_state.processed]
                sel_labels = st.multiselect("Select invoices to update:", inv_labels)
            if st.button("🔄 Regenerate with new date", type="primary"):
                count = 0
                for i, item in enumerate(st.session_state.processed):
                    label = f"{item['filename']}  ({item['date']})"
                    if label in sel_labels:
                        try:
                            new_stamped = process_one(
                                item["original_bytes"], item["user"],
                                item["vendor"], item["cc"], item["gl"], new_date,
                            )
                            st.session_state.processed[i]["pdf_bytes"] = new_stamped
                            st.session_state.processed[i]["date"]      = new_date.strftime("%d/%m/%Y")
                            st.session_state.processed[i]["date_obj"]  = new_date
                            count += 1
                        except Exception as e:
                            st.error(f"Error in {item['filename']}: {e}")
                if count:
                    st.success(f"✅ {count} invoice(s) regenerated with date {new_date.strftime('%d/%m/%Y')}")
                    st.rerun()

        st.divider()

        to_delete = []
        for i, item in enumerate(st.session_state.processed):
            col1, col2, col3, col4 = st.columns([4, 1.5, 1, 1])
            with col1:
                st.markdown(f"📄 **{item['filename']}**")
                st.caption(
                    f"Vendor: `{item['vendor']}` | CC: `{item['cc']}` | "
                    f"GL: `{item['gl']}` | By: **{item['user']}** | "
                    f"Date: `{item['date']}` | Processed: {item['ts']}"
                )
            with col2:
                st.download_button(
                    "⬇️ Download",
                    data=item["pdf_bytes"],
                    file_name=item["filename"],
                    mime="application/pdf",
                    key=f"dl_{i}",
                    use_container_width=True,
                )
            with col3:
                if st.button("🗑️", key=f"del_{i}", help="Remove from list"):
                    to_delete.append(i)
            with col4:
                st.caption(f"#{i+1}")
            st.markdown("<hr style='margin:6px 0; border-color:#eee'>", unsafe_allow_html=True)

        if to_delete:
            for i in sorted(to_delete, reverse=True):
                st.session_state.processed.pop(i)
            st.rerun()


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — DATABASE
# ══════════════════════════════════════════════════════════════════════════════
with tab_db:
    db_t1, db_t2 = st.tabs(["🏢 Vendors / Cost Centres", "📊 GL Accounts"])

    with db_t1:
        st.subheader("Vendor Table")
        st.caption("Maps the **prefix** (first 2 letters of the Order No.) → Vendor and Cost Centre")
        df_prov = pd.DataFrame(st.session_state.proveedores)
        edited_prov = st.data_editor(
            df_prov,
            column_config={
                "prefijo": st.column_config.TextColumn("Prefix", width="small",
                    help="First 2 letters of the customer Order No."),
                "vendor":  st.column_config.TextColumn("Vendor No.", width="medium"),
                "cc":      st.column_config.TextColumn("Cost Centre", width="medium"),
            },
            num_rows="dynamic",
            use_container_width=True,
            key="prov_editor",
        )
        if st.button("💾 Save changes — Vendors", type="primary"):
            st.session_state.proveedores = [
                r for r in edited_prov.to_dict("records") if r.get("prefijo")
            ]
            st.success("✅ Vendor table updated")

    with db_t2:
        st.subheader("GL Codes Table")
        st.caption("Maps the **product code** from the invoice → GL Account")
        df_gl = pd.DataFrame(st.session_state.gl_codes)
        edited_gl = st.data_editor(
            df_gl,
            column_config={
                "codigo": st.column_config.TextColumn("Product Code", width="medium"),
                "gl":     st.column_config.TextColumn("GL Account", width="medium"),
            },
            num_rows="dynamic",
            use_container_width=True,
            key="gl_editor",
        )
        if st.button("💾 Save changes — GL", type="primary"):
            st.session_state.gl_codes = [
                r for r in edited_gl.to_dict("records") if r.get("codigo")
            ]
            st.success("✅ GL table updated")

        st.divider()
        st.subheader("📥 Import from Excel (maestro_contable.xlsx)")
        xl_up = st.file_uploader("Upload Excel file", type=["xlsx"], key="xl_import")
        if xl_up:
            try:
                wb = openpyxl.load_workbook(xl_up)
                if "proveedores" in wb.sheetnames:
                    ws = wb["proveedores"]
                    rows = list(ws.iter_rows(min_row=2, values_only=True))
                    st.session_state.proveedores = [
                        {"prefijo": str(r[0]), "vendor": str(r[1]), "cc": str(r[2])}
                        for r in rows if r[0]
                    ]
                if "cuentas_gl" in wb.sheetnames:
                    ws = wb["cuentas_gl"]
                    rows = list(ws.iter_rows(min_row=2, values_only=True))
                    st.session_state.gl_codes = [
                        {"codigo": str(r[0]), "gl": str(r[1])}
                        for r in rows if r[0] and r[1]
                    ]
                st.success(
                    f"✅ Imported: {len(st.session_state.proveedores)} vendors, "
                    f"{len(st.session_state.gl_codes)} GL codes"
                )
                st.rerun()
            except Exception as e:
                st.error(f"Error importing Excel: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 6 — SETTINGS
# ══════════════════════════════════════════════════════════════════════════════
with tab_cfg:
    st.subheader("General Settings")
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**👥 System Users**")
        users_txt = st.text_area(
            "One user per line:",
            value="\n".join(st.session_state.usuarios),
            height=130,
        )
        if st.button("💾 Save users"):
            new_users = [u.strip() for u in users_txt.splitlines() if u.strip()]
            if new_users:
                st.session_state.usuarios = new_users
                st.success(f"✅ {len(new_users)} user(s) saved")
            else:
                st.warning("The list cannot be empty")

    with col2:
        st.markdown("**📍 Stamp Position on PDF**")
        st.caption("Invoice PDFs are landscape (792 × 612 pts). The stamp has a white background with a red border.")

        c_x, c_y = st.columns(2)
        with c_x:
            new_sx = st.number_input("X — from left", 0, 780, st.session_state.stamp_x, step=5)
        with c_y:
            new_sy = st.number_input("Y — stamp top (from bottom)", 0, 610, st.session_state.stamp_y_top, step=5)

        c_w, c_h = st.columns(2)
        with c_w:
            new_sw = st.number_input("Width", 80, 400, st.session_state.stamp_w, step=5)
        with c_h:
            new_sh = st.number_input("Height", 40, 200, st.session_state.stamp_h, step=5)

        if st.button("💾 Save position"):
            st.session_state.stamp_x     = new_sx
            st.session_state.stamp_y_top = new_sy
            st.session_state.stamp_w     = new_sw
            st.session_state.stamp_h     = new_sh
            st.success("✅ Position updated")

        st.info(
            "💡 **Default values (centred top):** X=281, Y=594, Width=230, Height=82  \n"
            "Adjust if the stamp does not land in the correct area of the invoice."
        )

    st.divider()

    with st.expander("ℹ️ System Information"):
        st.markdown(f"""
        **Atlantic Invoice Tools v3.0**

        **Workflow:**
        1. **✂️ Split** — Upload Atlantic batch PDFs → auto-detect invoices → split into individual files
           named `{{Invoice No}} {{CC}} {{BOL}}.pdf`
        2. **🔗 Match** — Upload split invoices + PO PDFs → merge matched pairs
        3. **📤 Codify** — Stamp invoices with Vendor / CC / GL / Date

        **Split naming logic:**
        - Invoice No: from `INVOICE No/No DE FACTURE` field
        - CC: first 2 chars of Customer Order No. → lookup in Vendor table → first 2 chars of `cc` field
          *(e.g. MD → EV01 → **EV**)*
        - BOL: second line of Bill of Lading column if two lines exist, otherwise the single line value

        **Coding logic:**
        - The first **2 characters** of the Customer Order No. determine the CC prefix
        - The prefix is looked up in the **Vendor table** → Vendor + Cost Centre
        - The **Product Code** is extracted from the invoice and looked up in the **GL table** → GL Account
        - **Type-6 invoice exception:** if the invoice No. starts with `6` →
          Vendor = `{VENDOR_EXCEPCION}` (fixed), CC = manual user selection

        **Vendors in database:** {len(st.session_state.proveedores)}
        **GL codes in database:** {len(st.session_state.gl_codes)}
        **Users:** {', '.join(st.session_state.usuarios)}
        """)
