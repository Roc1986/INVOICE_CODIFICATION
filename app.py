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
from datetime import date, datetime, timedelta
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
        "active_module":       None,
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
        # AP Audit state
        "audit_results":       None,
        "audit_data_count":    0,
        # Payment Packager state
        "payment_result":          None,
        "payment_batches":         [],   # [{"label", "files": [{"name", "bytes"}]}]
        "payment_batch_form_key":  0,
        # Reconciliation state
        "recon_results":           None,
        "recon_zip":               None,
        "recon_upload_key":        0,
        "recon_statement_total":   None,
        "recon_grand_total":       None,
        "recon_check_po":          True,
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


def _get_processed_zip() -> bytes:
    items = st.session_state.processed
    if not items:
        return b""
    sig = tuple(item["filename"] for item in items)
    if st.session_state.get("_proc_zip_sig") == sig and st.session_state.get("_proc_zip"):
        return st.session_state["_proc_zip"]
    z = make_zip(items)
    st.session_state["_proc_zip"]     = z
    st.session_state["_proc_zip_sig"] = sig
    return z


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
# ── COURU CODE FUNCTIONS ──────────────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────────────────────

def extract_couru_data(pdf_bytes: bytes, filename: str = "") -> dict:
    result = {
        "invoice_no": None, "date": None,
        "gl": None, "cc": None, "subtotal": None, "error": None,
    }
    try:
        with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
            if not pdf.pages:
                result["error"] = "Empty PDF"
                return result
            full_text = ""
            for page in pdf.pages:
                full_text += (page.extract_text() or "") + "\n"
            text_upper = full_text.upper()
            lines = [ln.strip() for ln in full_text.splitlines() if ln.strip()]

            # Invoice number
            for i, line in enumerate(lines):
                if "INVOICE NO" in line.upper() and "FACTURE" in line.upper():
                    m = re.search(r"(\d{7,10})\s*$", line)
                    if m:
                        result["invoice_no"] = m.group(1)
                        break
                    for j in range(max(0, i - 5), i):
                        if re.fullmatch(r"\d{6,10}", lines[j]):
                            result["invoice_no"] = lines[j]
                            break
                    break
            if not result["invoice_no"] and filename:
                m = re.match(r"(\d{6,10})", filename)
                if m:
                    result["invoice_no"] = m.group(1)

            # Invoice date — last date on the customer-order header line
            date_pat = r'(\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\s+\d{4})'
            for line in lines:
                if re.search(r'\d{6}\s+\d{2,3}\s+[A-Za-z]{2}\d+', line, re.IGNORECASE):
                    dates = re.findall(date_pat, line, re.IGNORECASE)
                    if dates:
                        result["date"] = dates[-1]
                        break
            if not result["date"]:
                m = re.search(date_pat, full_text, re.IGNORECASE)
                if m:
                    result["date"] = m.group(1)

            # CC from customer order prefix
            for line in lines:
                m = re.search(r"\d{6}\s+\d{2,3}\s+\|?\s*([A-Za-z]{2}\d{4,7})\b", line, re.IGNORECASE)
                if m:
                    _, cc = get_vendor_cc(m.group(1)[:2].upper())
                    result["cc"] = cc
                    break

            # GL from product codes
            known_codes = sorted(
                [str(r["codigo"]).upper() for r in st.session_state.gl_codes],
                key=len, reverse=True,
            )
            if known_codes:
                _pat = r"\b(?:" + "|".join(re.escape(c) for c in known_codes) + r")\b"
                _m = re.search(_pat, text_upper)
                if _m:
                    result["gl"] = get_gl(_m.group(0))

            # Subtotal (total before taxes) — find amount near "Sub Total" label
            for i, line in enumerate(lines):
                if re.search(r'\bsub\s*total\b', line, re.IGNORECASE):
                    amounts = re.findall(r'[\d,]+\.\d{2}', line)
                    if amounts:
                        result["subtotal"] = amounts[0].replace(",", "")
                        break
                    for j in range(max(0, i - 4), i):
                        amounts = re.findall(r'^([\d,]+\.\d{2})$', lines[j])
                        if amounts:
                            result["subtotal"] = amounts[0].replace(",", "")
                            break
                    if result["subtotal"]:
                        break
                    for j in range(i + 1, min(len(lines), i + 3)):
                        amounts = re.findall(r'([\d,]+\.\d{2})', lines[j])
                        if amounts:
                            result["subtotal"] = amounts[0].replace(",", "")
                            break
                    break

    except Exception as e:
        result["error"] = str(e)
    return result


# ─────────────────────────────────────────────────────────────────────────────
# ── AP AUDIT VALIDATION FUNCTIONS ─────────────────────────────────────────────
# ─────────────────────────────────────────────────────────────────────────────

def _parse_date_str(s: str) -> "date | None":
    """Parse a date string in multiple formats → date object."""
    if not s:
        return None
    s = s.strip()
    for fmt in (
        "%d %b %Y", "%d %B %Y",
        "%d-%b-%Y", "%d-%B-%Y",
        "%d-%b-%y", "%d-%B-%y",
        "%d/%m/%Y", "%m/%d/%Y",
        "%d/%m/%y", "%m/%d/%y",
        "%Y-%m-%d", "%Y/%m/%d",
        "%d.%m.%Y",
    ):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def calc_due_date(invoice_date) -> "date | None":
    """
    Quincena rule: invoice_date + 60 days → round UP to the 15th or 30th.
    The second quincena is always the 30th (never 31).
    """
    if isinstance(invoice_date, str):
        invoice_date = _parse_date_str(invoice_date)
    if not invoice_date:
        return None
    base = invoice_date + timedelta(days=60)
    if base.day <= 15:
        return date(base.year, base.month, 15)
    return date(base.year, base.month, 30)


def extract_invoice_date(pdf_bytes: bytes) -> "date | None":
    """
    Extract the INVOICE DATE from an Atlantic invoice PDF.
    Tries: 1) 'INVOICE DATE' label search  2) last date on the customer-order line.
    """
    date_pat = r'(\d{1,2}\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\w*\s+\d{4})'
    try:
        with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
            full_text = ""
            for page in pdf.pages:
                full_text += (page.extract_text() or "") + "\n"
            lines = [ln.strip() for ln in full_text.splitlines() if ln.strip()]

            # Method 1: find 'INVOICE DATE' label and grab the adjacent date
            for i, line in enumerate(lines):
                if re.search(r'INVOICE\s+DATE', line, re.IGNORECASE):
                    window = " ".join(lines[i:min(i + 3, len(lines))])
                    m = re.search(date_pat, window, re.IGNORECASE)
                    if m:
                        return _parse_date_str(m.group(1))

            # Method 2: last date on the customer-order header line
            for line in lines:
                if re.search(r'\d{6}\s+\d{2,3}\s+[A-Za-z]{2}\d+', line, re.IGNORECASE):
                    dates = re.findall(date_pat, line, re.IGNORECASE)
                    if dates:
                        return _parse_date_str(dates[-1])
    except Exception:
        pass
    return None


def parse_audit_report(pdf_bytes: bytes) -> dict:
    """
    Parse the A/P Voucher Audit Listing PDF (Crystal Reports format).

    All pages are merged into a single coordinate space (y-offset per page)
    so invoice blocks that span a page boundary are handled correctly.
    Each block runs from [APINV_y, next_APINV_y) — no cross-invoice contamination.
    Invoice date = date on the same line as the APINV/invoice-number row;
    due date = the next date found in the rows below it (yk+6..yk+45).
    Total is located by the 'Total:' label; falls back to max amount in block.
    """
    records = {}
    known_ccs = {r["cc"].upper() for r in st.session_state.proveedores}

    _ISO_PAT    = re.compile(r'^\d{4}-\d{2}-\d{2}$')
    _OTHER_PATS = [
        re.compile(r'^\d{1,2}-(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-\d{2,4}$', re.I),
        re.compile(r'^\d{2}/\d{2}/\d{4}$'),
        re.compile(r'^\d{4}/\d{2}/\d{2}$'),
        re.compile(r'^\d{2}\.\d{2}\.\d{4}$'),
    ]
    # Non-anchored versions for searching inside a reconstructed character run
    _DATE_SEARCH_PATS = [
        re.compile(r'\d{4}-\d{2}-\d{2}'),
        re.compile(r'\d{1,2}-(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)-\d{2,4}', re.I),
        re.compile(r'\d{2}/\d{2}/\d{4}'),
        re.compile(r'\d{4}/\d{2}/\d{2}'),
        re.compile(r'\d{2}\.\d{2}\.\d{4}'),
    ]

    def _is_date_word(text: str) -> "date | None":
        t = text.strip()
        if _ISO_PAT.match(t):
            return _parse_date_str(t)
        for p in _OTHER_PATS:
            if p.match(t):
                return _parse_date_str(t)
        return None

    def _row_date(ys: list) -> "date | None":
        """First date word found scanning the given rows top-to-bottom."""
        for y in ys:
            for w in sorted(row_map.get(y, []), key=lambda w: w["x0"]):
                d = _is_date_word(w["text"])
                if d:
                    return d
        return None

    def _chars_date(top_lo: float, top_hi: float, x_lo, x_hi) -> "date | None":
        """
        Read a date directly from the character stream within a y/x box.

        Crystal Reports can overflow a long Party name into the date
        column's x-range; both text runs then share the same x0 per
        character cell, and pdfplumber's word merge garbles them into one
        string (e.g. "0M2L6-08-12" instead of "2026-08-12"). Rebuilding
        the cell from characters and picking, at each x0, the glyph most
        likely to be a genuine date character (digit > separator > other)
        recovers the real date even when the overflow text itself
        contains a hyphen at that same position.
        """
        if x_lo is None or x_hi is None:
            return None
        cell = [c for c in all_chars
                if top_lo <= c["top"] <= top_hi and x_lo - 5 <= c["x0"] < x_hi]
        if not cell:
            return None

        def _priority(t: str) -> int:
            if t.isdigit():
                return 2
            if t in "-/.":
                return 1
            return 0

        by_x = {}
        for c in cell:
            x0 = round(c["x0"], 1)
            t = c["text"]
            pr = _priority(t)
            prev = by_x.get(x0)
            if prev is None or pr > prev[1]:
                by_x[x0] = (t, pr)
        s = "".join(v[0] for _, v in sorted(by_x.items()))
        for p in _DATE_SEARCH_PATS:
            m = p.search(s)
            if m:
                return _parse_date_str(m.group(0))
        return None

    try:
        # ── Merge all pages into one coordinate space ────────────────────────────
        all_words: list = []
        all_chars: list = []
        y_offset = 0.0
        with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                pw = page.extract_words(
                    x_tolerance=3, y_tolerance=3,
                    keep_blank_chars=False, use_text_flow=False,
                )
                for w in pw:
                    nw = dict(w)
                    nw["top"]    = w["top"]    + y_offset
                    nw["bottom"] = w["bottom"] + y_offset
                    all_words.append(nw)
                for ch in page.chars:
                    nc = dict(ch)
                    nc["top"] = ch["top"] + y_offset
                    all_chars.append(nc)
                y_offset += page.height

        if not all_words:
            return records

        # Date column x-range, from the "Invoice Date / Due Date" header
        # (both sub-labels start at the same x0 as the column itself).
        date_col_x0, date_col_x1 = None, None
        for w in all_words:
            if w["text"] in ("Invoice", "Due") and date_col_x0 is None:
                date_col_x0 = w["x0"]
            if w["text"] == "Year" and date_col_x1 is None:
                date_col_x1 = w["x0"]
            if date_col_x0 is not None and date_col_x1 is not None:
                break

        # Group into y-rows with 6 pt tolerance (Crystal Reports may offset
        # words on the same visual line by up to 5 pt)
        row_map: dict = {}
        for w in all_words:
            yk = round(w["top"] / 6) * 6
            row_map.setdefault(yk, []).append(w)

        sorted_ys = sorted(row_map.keys())

        # ── Pass 1: find all APINV rows ──────────────────────────────────────────
        apinv_rows: list = []  # [(yk, invoice_no, voucher_no)]
        for yk in sorted_ys:
            row_ws = row_map[yk]
            # Sort left-to-right so the regex finds "APINV <number>" in order
            row_ws_x = sorted(row_ws, key=lambda w: w["x0"])
            row_tx = " ".join(w["text"] for w in row_ws_x)
            if "APINV" not in row_tx.upper():
                continue

            # Voucher number: 1-4 digit integer immediately before "APINV"
            mv = re.search(r'(?:^|\s)(\d{1,4})\s+APINV\b', row_tx, re.IGNORECASE)
            voucher_no = mv.group(1) if mv else None

            # Regex on joined row text (works when APINV and number share a band)
            m = re.search(r'APINV\s+(\d{7,10})\b', row_tx, re.IGNORECASE)
            if m:
                apinv_rows.append((yk, m.group(1), voucher_no))
                continue

            # Fallback: find "APINV" word, then look right across nearby y-bands
            apinv_x1 = None
            for w in row_ws_x:
                if "APINV" in w["text"].upper():
                    apinv_x1 = w["x1"]
                    break
            if apinv_x1 is not None:
                nearby = [y for y in sorted_ys if abs(y - yk) <= 12]
                found = False
                for ny in nearby:
                    for w in sorted(row_map[ny], key=lambda w: w["x0"]):
                        if w["x0"] >= apinv_x1 - 5 and re.match(r'^\d{7,10}$', w["text"]):
                            apinv_rows.append((yk, w["text"], voucher_no))
                            found = True
                            break
                    if found:
                        break

        # ── Pass 2: extract fields from each bounded block ───────────────────────
        for idx, (yk, invoice_no, voucher_no) in enumerate(apinv_rows):
            next_yk = (apinv_rows[idx + 1][0]
                       if idx + 1 < len(apinv_rows)
                       else max(sorted_ys) + 200)

            block_ys    = [y for y in sorted_ys if yk <= y < next_yk]
            block_words = [w for y in block_ys for w in row_map[y]]
            block_text  = " ".join(w["text"] for w in block_words)

            # Vendor: 10 digits starting with 01
            vendor = None
            mv = re.search(r'\b(01\d{8})\b', block_text)
            if mv:
                vendor = mv.group(1)

            # GL: 6 digits starting with 3
            gl = None
            mg = re.search(r'\b(3\d{5})\b', block_text)
            if mg:
                gl = mg.group(1)

            # CC: known values only
            cc = None
            for mc in re.finditer(r'\b([A-Z]{2}\d{2})\b', block_text):
                if mc.group(1).upper() in known_ccs:
                    cc = mc.group(1).upper()
                    break

            # Invoice date: the date printed ON THE SAME LINE as the APINV /
            # invoice-number row (±6 pt — Crystal Reports row tolerance).
            # Due date: the next date word in the following rows within the
            # block (it sits directly under the invoice date, one line down).
            # Character-based read (handles Party-name overflow into the
            # date column) is tried first; word-based scan is the fallback.
            block_cap = min(yk + 45, next_yk)
            inv_date = _chars_date(yk - 6, yk + 6, date_col_x0, date_col_x1)
            due_date = _chars_date(yk + 6, block_cap, date_col_x0, date_col_x1)

            if inv_date is None:
                same_row_ys = [y for y in sorted_ys if abs(y - yk) <= 6]
                inv_date = _row_date(same_row_ys)
            if due_date is None:
                below_ys = [y for y in sorted_ys if yk + 6 < y <= block_cap]
                due_date = _row_date(below_ys)

            # Total: find the 'Total:' label row (not Sub-Total / Sous-Total)
            total_amt = None
            for y in block_ys:
                row_ws2 = row_map[y]
                row_tx2 = " ".join(w["text"] for w in row_ws2).upper()
                if (re.search(r'\bTOTAL\b', row_tx2)
                        and "SUB" not in row_tx2
                        and "SOUS" not in row_tx2):
                    for w in row_ws2:
                        clean = w["text"].replace(",", "")
                        if re.match(r'^\d+\.\d{2}$', clean):
                            val = float(clean)
                            if val >= 10.0:
                                total_amt = val
                                break
                    if total_amt is not None:
                        break

            # Fallback: largest amount ≥ 10 in the block
            if total_amt is None:
                raw_amounts = re.findall(r'\b([\d,]+\.\d{2})\b', block_text)
                amounts_num = sorted({
                    float(a.replace(",", "")) for a in raw_amounts
                    if float(a.replace(",", "")) >= 10.0
                })
                total_amt = amounts_num[-1] if amounts_num else None

            records[invoice_no] = {
                "voucher":      voucher_no,
                "invoice_date": inv_date,
                "due_date":     due_date,
                "cc":           cc,
                "vendor":       vendor,
                "gl":           gl,
                "total":        total_amt,
            }

    except Exception:
        pass

    return records


def extract_invoice_amounts(pdf_bytes: bytes) -> dict:
    """
    Extract net (subtotal), combined taxes, and total from an Atlantic invoice PDF.
    Returns { "net": str|None, "taxes": float|None, "total": str|None }
    """
    result = {"net": None, "taxes": None, "total": None}
    try:
        with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
            full_text = ""
            for page in pdf.pages:
                full_text += (page.extract_text() or "") + "\n"
        lines = [ln.strip() for ln in full_text.splitlines() if ln.strip()]

        taxes_acc = None
        last_total = None

        for line in lines:
            upper = line.upper()
            amounts = re.findall(r'([\d,]+\.\d{2})', line)
            if not amounts:
                continue
            val = amounts[0].replace(",", "")

            if re.search(r'\bSUB\s*TOTAL\b|\bSOUS.?TOTAL\b', upper):
                result["net"] = val

            elif re.search(r'\bGST\b|\bTPS\b|G\.S\.T|T\.P\.S', upper):
                taxes_acc = (taxes_acc or 0.0) + float(val)

            elif re.search(r'\bQST\b|\bTVQ\b|Q\.S\.T|T\.V\.Q', upper):
                taxes_acc = (taxes_acc or 0.0) + float(val)

            elif re.search(r'^\s*TOTAL\b', upper):
                last_total = val

        if taxes_acc is not None:
            result["taxes"] = round(taxes_acc, 2)
        if last_total:
            result["total"] = last_total

        # Fallback: compute missing value from the other two
        net_f   = float(result["net"])   if result["net"]   else None
        total_f = float(result["total"]) if result["total"] else None
        if net_f is not None and total_f is not None:
            computed_taxes = round(total_f - net_f, 2)
            # Use computed taxes if direct extraction is missing or clearly wrong (>1 CAD off)
            if result["taxes"] is None or abs(result["taxes"] - computed_taxes) > 1.0:
                result["taxes"] = computed_taxes

    except Exception:
        pass
    return result


# ─────────────────────────────────────────────────────────────────────────────
# ── STATEMENT RECONCILIATION FUNCTIONS ────────────────────────────────────────
# ─────────────────────────────────────────────────────────────────────────────

def _recon_parse_amount(raw) -> "float | None":
    """Parse an amount that may be a number already, or text like '1,234.56' /
    '1,234.56-' (Crystal Reports puts the minus sign for credits at the end)."""
    if raw is None:
        return None
    if isinstance(raw, (int, float)):
        return float(raw)
    s = str(raw).strip().replace(",", "")
    if not s:
        return None
    neg = s.endswith("-")
    if neg:
        s = s[:-1]
    try:
        val = float(s)
    except ValueError:
        return None
    return -val if neg else val


_RECON_GRAND_TOTAL_PAT = re.compile(
    r'Total\s+Ow[a-z]*\s*:?\s*(?:CAD)?\s*([\d,]+\.\d{2})', re.IGNORECASE
)


def _recon_find_grand_total(text: str) -> "float | None":
    """Find the statement's own 'Total Owing' figure in free text, so the UI
    can show it next to the sum of the parsed lines as a sanity check that
    the file was read correctly."""
    m = _RECON_GRAND_TOTAL_PAT.search(text)
    if not m:
        return None
    return _recon_parse_amount(m.group(1))


def parse_atlantic_statement_pdf(pdf_bytes: bytes) -> tuple:
    """
    Parse an Atlantic 'Statement of Account' PDF.
    Each line looks like:
      24-03-26 WPD 87059502 R0015751 37,303.43 37,303.43 143
    (date, type code, invoice #, customer reference / PO, invoice amount,
    balance owing, days). Credit-memo lines carry a trailing '-' on the amount.
    Returns (records, grand_total) — grand_total is the 'Total Owing' figure
    printed on the statement itself, or None if it couldn't be found.
    """
    records = []
    grand_total = None
    line_pat = re.compile(
        r'^(\d{2}-\d{2}-\d{2})\s+([A-Z]{3})\s+(\d{6,9})\s+(\S+)\s+'
        r'([\d,]+\.\d{2}-?)\s+([\d,]+\.\d{2}-?)\s+(\d+)\s*$'
    )
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if grand_total is None:
                grand_total = _recon_find_grand_total(text)
            for line in text.splitlines():
                m = line_pat.match(line.strip())
                if not m:
                    continue
                date_str, typ, inv_no, ref, amount_str, balance_str, day = m.groups()
                try:
                    inv_date = datetime.strptime(date_str, "%d-%m-%y").date()
                except ValueError:
                    inv_date = None
                records.append({
                    "invoice_no": inv_no,
                    "type":       typ,
                    "po_ref":     ref,
                    "invoice_date": inv_date,
                    "amount":     _recon_parse_amount(amount_str),
                    "balance":    _recon_parse_amount(balance_str),
                    "day":        int(day),
                })
    return records, grand_total


def parse_atlantic_statement_excel(file_bytes: bytes) -> tuple:
    """
    Parse the Excel export of the same Atlantic statement (B2Win 'b2win'
    sheet). Columns: CAD marker | date | type | invoice # | cust reference |
    invoice amount | balance owing | day.
    Returns (records, grand_total) — see parse_atlantic_statement_pdf.
    """
    records = []
    grand_total = None
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    ws = wb.active
    for sheet in wb.worksheets:
        if "b2win" in sheet.title.lower():
            ws = sheet
            break

    inv_pat = re.compile(r'^\d{6,9}$')
    for row in ws.iter_rows(values_only=True):
        if not row or len(row) < 8:
            continue
        marker, inv_date, typ, inv_no, ref, amount, balance, day = row[:8]
        if str(marker).strip().upper() != "CAD":
            continue
        if grand_total is None and any("Total" in str(c) for c in row if c):
            # Cells split words arbitrarily (e.g. 'Total Ow' | 'ing ') — join
            # with no separator so the word and the trailing amount line up.
            row_text = "".join(str(c) for c in row if c)
            grand_total = _recon_find_grand_total(row_text)
        if not inv_no or not inv_pat.match(str(inv_no).strip()):
            continue
        if hasattr(inv_date, "date"):
            inv_date = inv_date.date()
        records.append({
            "invoice_no": str(inv_no).strip(),
            "type":       str(typ).strip() if typ else None,
            "po_ref":     str(ref).strip() if ref else None,
            "invoice_date": inv_date if isinstance(inv_date, date) else None,
            "amount":     _recon_parse_amount(amount),
            "balance":    _recon_parse_amount(balance),
            "day":        int(day) if day is not None else None,
        })
    return records, grand_total


def parse_atlantic_statement(filename: str, file_bytes: bytes) -> tuple:
    """Dispatch to the PDF or Excel parser based on the file extension.
    Returns (records, grand_total)."""
    if filename.lower().endswith(".pdf"):
        return parse_atlantic_statement_pdf(file_bytes)
    return parse_atlantic_statement_excel(file_bytes)


def _recon_parse_amount_eu(raw) -> "float | None":
    """Parse a European-style amount — space thousands separator, comma
    decimal (Transport Bourret's statement format), e.g. '1 542,81' ->
    1542.81."""
    if raw is None:
        return None
    s = re.sub(r'\s+', '', str(raw)).replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return None


def parse_bourret_statement_pdf(pdf_bytes: bytes) -> tuple:
    """
    Parse a Transport Bourret 'État de compte / Statement' PDF.
    Each line looks like:
      07/21/2026 08/20/2026 14597313 2009575/MDBL236811 1 542,81$
    (invoice date, due date, invoice #, customer reference — often blank,
    sometimes a composite of several tokens — and amount). The reference is
    captured for display only; Bourret's reference is a shipment/BOL number,
    not a PO, so it isn't compared against anything (see check_po in
    RECON_VENDORS).

    Parsed by tokens rather than a single regex: the amount is always the
    last token (ends in ',DD$'); a plain 1-3 digit token right before it is
    a space-grouped thousands prefix ('1 542,81$' -> amount, not part of the
    reference) — everything else between the invoice number and the amount
    is the reference, however many tokens or separators (/, -) it contains.
    A pure regex can't tell a 6-digit numeric reference like '360676' apart
    from a thousands-prefix + amount ('360' + '676 946,01$' both look like
    valid amount shapes) — the token approach resolves it by only ever
    treating a *short* (<=3 digit) trailing token as a thousands group.
    """
    records = []
    grand_total = None
    date_pat        = re.compile(r'^\d{2}/\d{2}/\d{4}$')
    invoice_no_pat  = re.compile(r'^\d{7,9}$')
    amount_end_pat  = re.compile(r'^(\d+),(\d{2})\$$')
    thousands_pat   = re.compile(r'^\d{1,3}$')
    total_pat       = re.compile(r'Total\s+(\d{1,3}(?:\s\d{3})*,\d{2})\$')

    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if grand_total is None:
                m = total_pat.search(text)
                if m:
                    grand_total = _recon_parse_amount_eu(m.group(1))
            for line in text.splitlines():
                tokens = line.strip().split()
                if (len(tokens) < 4
                        or not date_pat.match(tokens[0])
                        or not date_pat.match(tokens[1])
                        or not invoice_no_pat.match(tokens[2])):
                    continue
                date_str, inv_no, rest = tokens[0], tokens[2], tokens[3:]
                if not rest or not amount_end_pat.match(rest[-1]):
                    continue
                amount_tokens = [rest[-1]]
                ref_tokens = rest[:-1]
                if ref_tokens and thousands_pat.match(ref_tokens[-1]):
                    amount_tokens.insert(0, ref_tokens.pop())
                try:
                    inv_date = datetime.strptime(date_str, "%m/%d/%Y").date()
                except ValueError:
                    inv_date = None
                records.append({
                    "invoice_no": inv_no,
                    "type":       None,
                    "po_ref":     " ".join(ref_tokens) or None,
                    "invoice_date": inv_date,
                    "amount":     _recon_parse_amount_eu(" ".join(amount_tokens).rstrip("$")),
                    "balance":    None,
                    "day":        None,
                })
    return records, grand_total


def parse_bourret_statement(filename: str, file_bytes: bytes) -> tuple:
    """Dispatch for Transport Bourret's statement. Only PDF has been seen
    from this vendor so far."""
    if filename.lower().endswith(".pdf"):
        return parse_bourret_statement_pdf(file_bytes)
    raise ValueError("Only PDF statements are supported for Transport Bourret right now.")


def parse_proden_statement_pdf(pdf_bytes: bytes) -> tuple:
    """
    Parse a Les Entreprises Proden 'État de compte' aging PDF.
    Each line looks like:
      Apr-20-26 305912 MLQD82865 87 / 117 4,799.06
    (invoice date, invoice #, PO#, a "days overdue / days since invoice"
    pair we don't need, and the outstanding amount). The amount is printed
    in whichever of the four aging-bucket columns applies (Courant / 31-60
    / 61-90 / >90 jours), so exactly one amount value appears per line no
    matter which bucket it's in — we don't need to track which column it
    came from for reconciliation purposes.
    The footer totals row has 5 numbers: a grand total followed by the four
    bucket subtotals; the grand total is used for the statement-total
    sanity check.
    """
    records = []
    grand_total = None
    line_pat = re.compile(
        r'^([A-Za-z]{3}-\d{2}-\d{2})\s+(\d{5,7})\s+(\S+)\s+\d+\s*/\s*\d+\s+([\d,]+\.\d{2})\s*$'
    )
    total_pat = re.compile(
        r'^([\d,]+\.\d{2})\s+[\d,]+\.\d{2}\s+[\d,]+\.\d{2}\s+[\d,]+\.\d{2}\s+[\d,]+\.\d{2}\s*$'
    )
    with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            for line in text.splitlines():
                line = line.strip()
                if grand_total is None:
                    m_tot = total_pat.match(line)
                    if m_tot:
                        grand_total = _recon_parse_amount(m_tot.group(1))
                        continue
                m = line_pat.match(line)
                if not m:
                    continue
                date_str, inv_no, po_ref, amount_str = m.groups()
                try:
                    inv_date = datetime.strptime(date_str, "%b-%d-%y").date()
                except ValueError:
                    inv_date = None
                records.append({
                    "invoice_no": inv_no,
                    "type":       None,
                    "po_ref":     po_ref,
                    "invoice_date": inv_date,
                    "amount":     _recon_parse_amount(amount_str),
                    "balance":    None,
                    "day":        None,
                })
    return records, grand_total


def parse_proden_statement(filename: str, file_bytes: bytes) -> tuple:
    """Dispatch for Proden's statement. Only PDF has been seen from this
    vendor so far."""
    if filename.lower().endswith(".pdf"):
        return parse_proden_statement_pdf(file_bytes)
    raise ValueError("Only PDF statements are supported for Proden right now.")


def parse_system_extract(file_bytes: bytes) -> dict:
    """
    Parse the accounting-system AP extract (any vendor — the export is the
    same GL/voucher layout regardless of who the invoices belong to).
    Returns { invoice_no: {"total", "costctr", "paid", "payment_date"} }.
    An invoice can span several GL allocation lines (same invoice #, same
    total, different account/cost-centre) — those are collapsed into one
    record; a payment on ANY of its lines marks the invoice as paid.
    """
    df = pd.read_excel(BytesIO(file_bytes))
    cols = {c.strip().lower(): c for c in df.columns}

    def col(*names):
        for n in names:
            if n in cols:
                return cols[n]
        return None

    c_inv   = col("invoice")
    c_total = col("invoicetotal")
    c_cc    = col("costctr")
    c_pay   = col("paymentno")
    c_paydt = col("paymentdate")
    if c_inv is None or c_total is None:
        return {}

    records = {}
    for _, row in df.iterrows():
        inv_no = str(row[c_inv]).strip() if pd.notna(row[c_inv]) else ""
        if not inv_no:
            continue
        total = float(row[c_total]) if pd.notna(row[c_total]) else 0.0
        paid  = c_pay is not None and pd.notna(row[c_pay])
        rec = records.get(inv_no)
        if rec is None:
            rec = {
                "invoice_no":   inv_no,
                "total":        total,
                "costctr":      row[c_cc] if c_cc is not None and pd.notna(row[c_cc]) else None,
                "paid":         False,
                "payment_date": None,
            }
            records[inv_no] = rec
        if paid:
            rec["paid"] = True
            if c_paydt is not None and pd.notna(row[c_paydt]):
                rec["payment_date"] = row[c_paydt]
    return records


def parse_extraction_excel(file_bytes: bytes) -> dict:
    """
    Parse the invoice-copy tracking Excel (folder listing exported from the
    Atlantic invoices share). Returns { invoice_no: {"po", "location", "filename"} }.
    Expects columns named 'Name' (file name), 'Colonne2' (PO extracted from
    the file name) and 'Colonne4' (folder location relative to the share root).
    """
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return {}
    header = [str(h).strip() if h else "" for h in rows[0]]

    def idx(*names):
        for n in names:
            if n in header:
                return header.index(n)
        return None

    idx_name = idx("Name")
    idx_po   = idx("Colonne2")
    idx_loc  = idx("Colonne4")
    if idx_name is None:
        return {}

    inv_pat = re.compile(r'^(\d{6,9})\b')
    records = {}
    for row in rows[1:]:
        name = row[idx_name] if idx_name < len(row) else None
        if not name:
            continue
        m = inv_pat.match(str(name).strip())
        if not m:
            continue
        inv_no = m.group(1)
        po  = row[idx_po]  if idx_po  is not None and idx_po  < len(row) else None
        loc = row[idx_loc] if idx_loc is not None and idx_loc < len(row) else None
        records[inv_no] = {
            "po":       str(po).strip() if po else None,
            "location": str(loc).strip() if loc else None,
            "filename": str(name).strip(),
        }
    return records


def reconcile_statement(statement: list, system: dict, extraction: dict, check_po: bool = True) -> dict:
    """
    Drive the reconciliation off the vendor statement (per the workflow: the
    statement is what the vendor sends us, so every line in it must be
    accounted for). For each statement line:
      1. Look it up in the copy-tracking extract by invoice number (this is
         independent of registration status — every entry gets a has_copy /
         extraction_po / location, even when the copy-tracking file wasn't
         supplied at all, in which case has_copy is simply False for all).
      2. Look it up in the system extract by invoice number.
         - Not found            -> pending_registration (has_copy tells you
           whether the file is on hand and just needs entering, or missing
           entirely — no separate top-level bucket for that distinction).
         - Found + amount matches -> matched
         - Found + amount differs -> amount_mismatch
      3. When check_po is True and there's a copy on file, compare its PO
         (from the file name) against the statement's PO/customer-reference
         -> po_mismatch (this can happen on a matched invoice too, not just
         a pending one). Some vendors' "reference" isn't a PO at all (e.g.
         Transport Bourret's is a shipment/BOL number) — check_po=False
         skips this comparison entirely for those.
    """
    buckets = {
        "matched": [], "amount_mismatch": [],
        "pending_registration": [], "po_mismatch": [],
    }
    for line in statement:
        inv_no = line["invoice_no"]
        entry = dict(line)

        ext_rec = extraction.get(inv_no)
        entry["has_copy"]      = ext_rec is not None
        entry["extraction_po"] = ext_rec["po"] if ext_rec else None
        entry["location"]      = ext_rec["location"] if ext_rec else None
        entry["filename"]      = ext_rec["filename"] if ext_rec else None

        if check_po:
            po_a = str(line.get("po_ref") or "").strip().upper()
            po_b = str((ext_rec or {}).get("po") or "").strip().upper()
            entry["po_mismatch"] = bool(ext_rec is not None and po_a and po_b and po_a != po_b)
        else:
            entry["po_mismatch"] = False

        sys_rec = system.get(inv_no)
        if sys_rec is not None:
            entry["registered"]   = True
            entry["system_total"] = sys_rec["total"]
            entry["costctr"]      = sys_rec["costctr"]
            entry["paid"]         = sys_rec["paid"]
            entry["payment_date"] = sys_rec["payment_date"]
            if abs(abs(line["amount"] or 0) - abs(sys_rec["total"] or 0)) > 0.02:
                buckets["amount_mismatch"].append(entry)
            else:
                buckets["matched"].append(entry)
        else:
            entry["registered"] = False
            buckets["pending_registration"].append(entry)

        if entry["po_mismatch"]:
            buckets["po_mismatch"].append(entry)

    return buckets


def build_recon_summary_rows(buckets: dict, statement_total: float, check_po: bool = True) -> list:
    """
    Build the reconciliation summary as a flat, hierarchical list of
    (label, count, amount) rows — mirroring the AP team's own manual
    reconciliation layout: a top-level Total / Reconciled / Not Reconciled
    split with a balancing check, then a breakdown of Not Reconciled by
    whether we have a copy on file, then a 'To Review' section for the
    cross-cutting discrepancy flags. Blank label = section break; None in
    count/amount = leave that cell empty. The PO Mismatch line is omitted
    when check_po is False (vendors whose "reference" isn't a PO).
    """
    matched   = buckets["matched"]
    mismatch  = buckets["amount_mismatch"]
    pending   = buckets["pending_registration"]
    po_mism   = buckets["po_mismatch"]
    pending_with_copy = [r for r in pending if r["has_copy"]]
    pending_no_copy   = [r for r in pending if not r["has_copy"]]

    def amt(rows):
        return round(sum((r.get("amount") or 0) for r in rows), 2)

    mismatch_amt = amt(mismatch)
    # "Reconciled" money-wise means "accounted for in the system" — so its
    # amount includes invoices registered with a discrepancy too (those are
    # still in the system, just flagged in "To Review" for the exact figure).
    # Its count, however, is the clean-match count only, so the discrepancy
    # count is visible on its own line in "To Review" instead of hiding
    # inside "Reconciled". Not Reconciled is simpler: not in the system at
    # all, so count and amount both mean the same population there.
    reconciled_amt      = round(amt(matched) + mismatch_amt, 2)
    not_reconciled_amt  = amt(pending)
    check = round(statement_total - reconciled_amt - not_reconciled_amt, 2)
    check = 0.0 if abs(check) < 0.005 else check

    rows = [
        ("Total Statement",                    len(matched) + len(pending) + len(mismatch), round(statement_total, 2)),
        ("Reconciled",                         len(matched), reconciled_amt),
        ("Not Reconciled",                     len(pending), not_reconciled_amt),
        ("Reconciliation Check",               None, check),
        ("", None, None),
        ("Not Reconciled — breakdown",         None, None),
        ("  Pending Registration (has copy)",  len(pending_with_copy), amt(pending_with_copy)),
        ("  No Copy on File",                  len(pending_no_copy), amt(pending_no_copy)),
        ("", None, None),
        ("To Review",                          None, None),
        ("  Amount Mismatch",                  len(mismatch), mismatch_amt),
    ]
    if check_po:
        rows.append(("  PO Mismatch", len(po_mism), amt(po_mism)))
    return rows


def make_recon_report(buckets: dict, statement_total: float, grand_total: "float | None",
                       check_po: bool = True) -> bytes:
    """Build the downloadable Excel report: a Summary sheet (hierarchical
    counts/totals plus the statement-total sanity check) plus one sheet per
    category. The PO Mismatch sheet is omitted when check_po is False."""
    sheet_specs = [
        ("Matched",              buckets["matched"]),
        ("Amount Mismatch",      buckets["amount_mismatch"]),
        ("Pending Registration", buckets["pending_registration"]),
    ]
    if check_po:
        sheet_specs.append(("PO Mismatch", buckets["po_mismatch"]))
    cols = [
        "invoice_no", "type", "invoice_date", "po_ref", "amount", "day",
        "registered", "system_total", "costctr", "paid", "payment_date",
        "has_copy", "extraction_po", "location", "po_mismatch", "filename",
    ]
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        summary_rows = [
            {"Category": label, "Invoice Count": count, "Total Amount (CAD)": amount}
            for label, count, amount in build_recon_summary_rows(buckets, statement_total, check_po)
        ]
        summary_rows.append({"Category": "", "Invoice Count": None, "Total Amount (CAD)": None})
        summary_rows.append({
            "Category": "Statement Total (per file)" if grand_total is not None
                        else "Statement Total (not found in file)",
            "Invoice Count": None,
            "Total Amount (CAD)": grand_total,
        })
        summary_rows.append({
            "Category": "Sum of Parsed Invoice Lines",
            "Invoice Count": None,
            "Total Amount (CAD)": round(statement_total, 2),
        })
        pd.DataFrame(summary_rows).to_excel(writer, sheet_name="Summary", index=False)
        for label, rows in sheet_specs:
            if rows:
                df = pd.DataFrame(rows)
                df = df[[c for c in cols if c in df.columns]]
            else:
                df = pd.DataFrame(columns=cols)
            df.to_excel(writer, sheet_name=label[:31], index=False)
    buf.seek(0)
    return buf.read()


# Registry of vendors the Reconciliation module can parse a statement for.
# The system extract and the copy-tracking Excel are vendor-agnostic (the
# user exports them already scoped to one vendor), only the statement needs
# a vendor-specific parser — add new vendors here as their templates are built.
# check_po: whether the statement's reference column is a PO worth comparing
# against the copy-tracking file name (False for vendors like Transport
# Bourret, whose reference is a shipment/BOL number, not a PO).
RECON_VENDORS = {
    "atlantic": {"label": "Atlantic Packaging",  "parse_statement": parse_atlantic_statement, "check_po": True},
    "bourret":  {"label": "Transport Bourret",   "parse_statement": parse_bourret_statement,  "check_po": False},
    "proden":   {"label": "Les Entreprises Proden", "parse_statement": parse_proden_statement, "check_po": True},
}


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
.module-card [data-testid="stVerticalBlockBorderWrapper"] {
    transition: box-shadow 0.15s ease, transform 0.15s ease;
}
.module-card:hover [data-testid="stVerticalBlockBorderWrapper"] {
    box-shadow: 0 4px 14px rgba(0,0,0,0.12);
    transform: translateY(-2px);
}
.module-icon { font-size: 34px; line-height: 1; }
.module-label { font-size: 17px; font-weight: 700; margin-top: 6px; }
.module-desc  { font-size: 13px; color: #6c757d; margin-top: 2px; min-height: 34px; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# MODULE REGISTRY
# ─────────────────────────────────────────────────────────────────────────────
MODULES = [
    {"key": "splitter", "icon": "✂️",  "label": "Invoice Splitter",  "desc": "Split batch PDFs into individual invoice files."},
    {"key": "matcher",  "icon": "🔗", "label": "Invoice Matcher",   "desc": "Match invoices with POs and merge into one PDF."},
    {"key": "coding",   "icon": "🏷️",  "label": "Invoice Coding",    "desc": "Stamp GL / Cost Centre codes on invoices."},
    {"key": "couru",    "icon": "📊", "label": "Couru Code",        "desc": "Extract coding data for Couru entry."},
    {"key": "audit",    "icon": "🔍", "label": "AP Audit",          "desc": "Validate the A/P voucher audit listing."},
    {"key": "recon",    "icon": "🧮", "label": "Statement Reconciliation", "desc": "Reconcile Atlantic's statement against the system and invoice copies."},
    {"key": "payment",  "icon": "💳", "label": "Payment Packager",  "desc": "Bundle invoices into payment batches."},
    {"key": "database", "icon": "🗄️",  "label": "Database",          "desc": "Manage vendors, GL codes & users."},
    {"key": "settings", "icon": "⚙️",  "label": "Settings",          "desc": "Configure the coding stamp position."},
]

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 📂 Modules")
    nav_labels = ["🏠 Home"] + [f"{m['icon']} {m['label']}" for m in MODULES]
    nav_keys   = [None] + [m["key"] for m in MODULES]
    nav_idx    = nav_keys.index(st.session_state.active_module) \
        if st.session_state.active_module in nav_keys else 0
    picked = st.selectbox("Go to", nav_labels, index=nav_idx, label_visibility="collapsed")
    picked_key = nav_keys[nav_labels.index(picked)]
    if picked_key != st.session_state.active_module:
        st.session_state.active_module = picked_key
        st.rerun()

    st.divider()
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
        zip_bytes = _get_processed_zip()
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
active_module = st.session_state.active_module

if active_module is None:
    st.markdown("""
    <h1 style='margin-bottom:0'>📄 Atlantic — Invoice Tools</h1>
    <p style='color:gray;margin-top:4px'>Pick a module to get started</p>
    """, unsafe_allow_html=True)

    n_cols = 4
    rows = [MODULES[i:i + n_cols] for i in range(0, len(MODULES), n_cols)]
    for row in rows:
        cols = st.columns(n_cols)
        for col, mod in zip(cols, row):
            with col:
                st.markdown('<div class="module-card">', unsafe_allow_html=True)
                with st.container(border=True):
                    st.markdown(
                        f"<div class='module-icon'>{mod['icon']}</div>"
                        f"<div class='module-label'>{mod['label']}</div>"
                        f"<div class='module-desc'>{mod['desc']}</div>",
                        unsafe_allow_html=True,
                    )
                    if st.button("Open →", key=f"open_{mod['key']}", use_container_width=True):
                        st.session_state.active_module = mod["key"]
                        st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)
else:
    active_mod = next(m for m in MODULES if m["key"] == active_module)
    col_home, col_title = st.columns([1, 6])
    with col_home:
        if st.button("🏠 Home", use_container_width=True):
            st.session_state.active_module = None
            st.rerun()
    with col_title:
        st.markdown(
            f"<h2 style='margin:2px 0 0 0'>{active_mod['icon']} {active_mod['label']}</h2>",
            unsafe_allow_html=True,
        )
    st.divider()

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — INVOICE SPLITTER
# ══════════════════════════════════════════════════════════════════════════════
if active_module == "splitter":
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
        for row_idx, item in enumerate(results):
            row_cls = "split-row split-warn" if item["warning"] else "split-row"
            icon    = "⚠️" if item["warning"] else "✅"

            col_info, col_dl = st.columns([5, 1])
            with col_info:
                pages_str = ", ".join(str(p) for p in item["source_pages"])
                page_lbl  = f"{item['page_count']} page{'s' if item['page_count'] > 1 else ''}"
                warn_html = (
                    f"<br><span style='color:#e08000;font-size:12px'>{item['warning']}</span>"
                    if item["warning"] else ""
                )
                st.markdown(
                    f"<div class='{row_cls}'>"
                    f"{icon} <b>{item['filename']}</b>"
                    f"&nbsp;&nbsp;|&nbsp;&nbsp;"
                    f"Invoice: <code>{item['invoice_no']}</code> &nbsp; "
                    f"CC: <code>{item['cc']}</code> &nbsp; "
                    f"BOL: <code>{item['bol']}</code> &nbsp; "
                    f"· {page_lbl} (source p. {pages_str})"
                    f"{warn_html}"
                    f"</div>",
                    unsafe_allow_html=True,
                )
            with col_dl:
                st.download_button(
                    "⬇️",
                    data=item["pdf_bytes"],
                    file_name=item["filename"],
                    mime="application/pdf",
                    key=f"spdl_{row_idx}",
                    use_container_width=True,
                    help=f"Download {item['filename']}",
                )

    elif not split_uploads:
        st.info("📂 Upload one or more Atlantic batch PDFs to get started.")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — INVOICE MATCHER
# ══════════════════════════════════════════════════════════════════════════════
if active_module == "matcher":
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
# TAB 3 — INVOICE CODING (upload + review + results unified)
# ══════════════════════════════════════════════════════════════════════════════
if active_module == "coding":
    if "upload_key" not in st.session_state:
        st.session_state.upload_key = 0

    # ── Top toolbar: upload + clear all ──────────────────────────────────────
    col_up, col_clear_upload, col_clear_all = st.columns([5, 1, 1])
    with col_up:
        uploaded = st.file_uploader(
            "Drag or select one or more PDFs",
            type=["pdf"],
            accept_multiple_files=True,
            help="Supports bulk upload. Only the first page of each invoice is stamped.",
            key=f"uploader_{st.session_state.upload_key}",
        )
    with col_clear_upload:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🗑️ Clear\nfiles", use_container_width=True,
                     help="Remove uploaded files to load a new batch"):
            st.session_state.upload_key += 1
            st.rerun()
    with col_clear_all:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🗑️ Clear\nall", use_container_width=True,
                     help="Clear uploaded files AND all coded results"):
            st.session_state.processed    = []
            st.session_state.upload_key  += 1
            st.rerun()

    # ── Upload / review section ───────────────────────────────────────────────
    if not uploaded:
        st.session_state.pop("_inv_ui_cache", None)
        st.session_state.pop("_inv_ui_sig", None)
        if not st.session_state.processed:
            st.info("📂 Upload invoices to get started. You can select multiple files at once.")
    else:
        st.success(f"✅ **{len(uploaded)} file(s)** loaded — analyzing…")
        st.divider()

        # Re-parse PDFs only when files OR the GL/vendor database changes
        _sig = (
            tuple((f.name, f.size) for f in uploaded),
            tuple(r["codigo"] for r in st.session_state.gl_codes),
            tuple(r["prefijo"] for r in st.session_state.proveedores),
        )
        if _sig != st.session_state.get("_inv_ui_sig"):
            _inv_ui = []
            for f in uploaded:
                raw = f.read()
                data = extract_invoice_data(raw, f.name)
                data["filename"] = f.name
                data["raw_bytes"] = raw
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
                _inv_ui.append(data)
            st.session_state["_inv_ui_cache"] = _inv_ui
            st.session_state["_inv_ui_sig"]   = _sig
        invoices_ui = st.session_state["_inv_ui_cache"]

        resolved_cc     = {}
        resolved_vendor = {}
        resolved_gl     = {}

        n_auto   = sum(
            1 for inv in invoices_ui
            if not inv.get("error")
            and (inv.get("invoice_no") or inv.get("customer_order"))
            and not inv["needs_cc"]
            and inv.get("gl_auto")
        )
        n_review = len(invoices_ui) - n_auto

        if n_auto:
            st.success(f"✅ {n_auto} invoice(s) auto-coded — no review needed.")
        if n_review:
            st.warning(f"⚠️ {n_review} invoice(s) require manual input — review below before coding.")

        for idx, inv in enumerate(invoices_ui):
            has_err    = inv.get("error") or (not inv.get("invoice_no") and not inv.get("customer_order"))
            needs_input = inv["needs_cc"] or not inv.get("gl_auto")
            auto_coded  = not has_err and not needs_input

            if auto_coded:
                # Pre-populate resolved dicts; no UI needed
                resolved_vendor[idx] = VENDOR_EXCEPCION if inv.get("is_six") else inv["vendor_auto"]
                resolved_cc[idx]     = inv["cc_auto"]
                resolved_gl[idx]     = inv["gl_auto"]
                continue

            icon = "❌" if has_err else "⚠️"
            with st.expander(f"{icon}  {inv['filename']}", expanded=True):
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

                cc_prev   = resolved_cc.get(idx, "??")
                gl_prev   = resolved_gl.get(idx, "??")
                vd_prev   = resolved_vendor.get(idx, "??")
                usr_prev  = current_user or "???"
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
                "🚀 Code Invoices",
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
                fname  = inv["filename"]
                progress.progress((idx + 1) / len(invoices_ui), text=f"Coding {fname}…")
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
                st.success(f"🎉 **{len(invoices_ui)} invoice(s)** coded successfully.")
                st.balloons()

    # ── Results section (always visible when there are coded invoices) ────────
    if st.session_state.processed:
        n = len(st.session_state.processed)
        st.divider()
        col_hdr, col_zip, col_del = st.columns([3, 2, 1])
        with col_hdr:
            st.subheader(f"📋 Coded Invoices — {n} file(s)")
        with col_zip:
            zip_all = _get_processed_zip()
            st.download_button(
                f"⬇️ Download ZIP ({n})",
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
                    f"Date: `{item['date']}` | Coded: {item['ts']}"
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
# TAB 4 — COURU CODE
# ══════════════════════════════════════════════════════════════════════════════
if active_module == "couru":
    st.subheader("📊 Couru Code — Invoice Report")
    st.markdown(
        "Upload invoice PDFs to extract key data and download an Excel report "
        "with invoice number, date, GL account, cost centre, and subtotal before taxes."
    )

    if "couru_upload_key" not in st.session_state:
        st.session_state.couru_upload_key = 0

    col_up, col_clear = st.columns([5, 1])
    with col_up:
        couru_files = st.file_uploader(
            "Drag or select one or more invoice PDFs",
            type=["pdf"],
            accept_multiple_files=True,
            key=f"couru_uploader_{st.session_state.couru_upload_key}",
        )
    with col_clear:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🗑️ Clear", use_container_width=True, key="couru_clear"):
            st.session_state.couru_upload_key += 1
            st.session_state.pop("couru_results", None)
            st.rerun()

    if not couru_files:
        st.info("📂 Upload invoice PDFs to get started.")
    else:
        st.success(f"✅ **{len(couru_files)} file(s)** loaded")

        if st.button("📊 Generate Excel Report", type="primary", key="couru_run"):
            couru_data = []
            prog = st.progress(0, text="Extracting…")
            for fi, f in enumerate(couru_files):
                prog.progress((fi + 1) / len(couru_files), text=f"Processing {f.name}…")
                raw = f.read()
                d = extract_couru_data(raw, f.name)
                couru_data.append({
                    "Invoice No": d["invoice_no"] or "—",
                    "Date":       d["date"]       or "—",
                    "GL":         d["gl"]         or "—",
                    "CC":         d["cc"]         or "—",
                    "Subtotal":   d["subtotal"]   or "—",
                    "Error":      d["error"]      or "",
                })
            prog.progress(1.0, text="✅ Done")
            st.session_state["couru_results"] = couru_data

        if st.session_state.get("couru_results"):
            rows = st.session_state["couru_results"]
            df = pd.DataFrame(rows)

            n_ok   = sum(1 for r in rows if not r["Error"] and r["Invoice No"] != "—")
            n_warn = sum(1 for r in rows if r["Error"] or r["GL"] == "—" or r["CC"] == "—")

            c1, c2, c3 = st.columns(3)
            with c1:
                st.markdown(
                    f"<div class='stat-box'><div class='stat-num blue'>{len(rows)}</div>"
                    f"<div class='stat-lbl'>Invoices processed</div></div>",
                    unsafe_allow_html=True,
                )
            with c2:
                st.markdown(
                    f"<div class='stat-box'><div class='stat-num green'>{n_ok}</div>"
                    f"<div class='stat-lbl'>Complete</div></div>",
                    unsafe_allow_html=True,
                )
            with c3:
                st.markdown(
                    f"<div class='stat-box'><div class='stat-num amber'>{n_warn}</div>"
                    f"<div class='stat-lbl'>Need review</div></div>",
                    unsafe_allow_html=True,
                )

            st.markdown("<br>", unsafe_allow_html=True)
            st.dataframe(df, use_container_width=True, hide_index=True)

            buf = BytesIO()
            display_df = df.drop(columns=["Error"]) if all(r == "" for r in df["Error"]) else df
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                display_df.to_excel(writer, index=False)
            buf.seek(0)
            st.download_button(
                f"⬇️ Download Excel ({len(rows)} invoices)",
                data=buf.read(),
                file_name=f"couru_report_{date.today().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
            )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 5 — AP AUDIT VALIDATION
# ══════════════════════════════════════════════════════════════════════════════
if active_module == "audit":
    st.subheader("🔍 AP Audit Validation")
    st.markdown(
        "Cross-reference invoice PDFs against the **A/P Voucher Audit Listing** "
        "to verify invoice number, date, payment date, cost centre, vendor, and GL "
        "before the session is permanently posted."
    )

    col_inv_up, col_rpt_up = st.columns(2)
    with col_inv_up:
        st.markdown("**📄 Invoice PDFs**")
        audit_inv_files = st.file_uploader(
            "Upload invoices",
            type=["pdf"],
            accept_multiple_files=True,
            key="audit_inv_upload",
            label_visibility="collapsed",
        )
        if audit_inv_files:
            st.caption(f"{len(audit_inv_files)} invoice(s) loaded")

    with col_rpt_up:
        st.markdown("**📋 A/P Voucher Audit Listing (PDF)**")
        audit_report_file = st.file_uploader(
            "Upload audit report",
            type=["pdf"],
            accept_multiple_files=False,
            key="audit_report_upload",
            label_visibility="collapsed",
        )
        if audit_report_file:
            st.caption(f"Report: **{audit_report_file.name}**")

    col_run_a, col_clr_a, _ = st.columns([2, 2, 5])
    with col_run_a:
        run_audit = st.button(
            "🔍 Run Validation",
            type="primary",
            use_container_width=True,
            disabled=not (audit_inv_files and audit_report_file),
        )
    with col_clr_a:
        if st.button("🗑️ Clear results", use_container_width=True, key="audit_clear"):
            st.session_state["audit_results"]    = None
            st.session_state["audit_data_count"] = 0
            st.rerun()

    if run_audit and audit_inv_files and audit_report_file:
        audit_data = parse_audit_report(audit_report_file.read())
        st.session_state["audit_data_count"] = len(audit_data)

        if not audit_data:
            st.warning(
                "⚠️ No records could be parsed from the audit report PDF. "
                "Please verify the file is the correct A/P Voucher Audit Listing."
            )
        else:
            def _cmp(a, b):
                if a is None or b is None:
                    return None
                if isinstance(a, date) and isinstance(b, date):
                    return a == b
                return str(a).strip().upper() == str(b).strip().upper()

            def _cmp_amt(a, b, tol=0.05):
                """Compare two floats with tolerance; None on either side → None."""
                if a is None or b is None:
                    return None
                try:
                    return abs(float(a) - float(b)) <= tol
                except (TypeError, ValueError):
                    return None

            val_results = []
            prog = st.progress(0, text="Validating invoices…")
            for fi, f in enumerate(audit_inv_files):
                prog.progress((fi + 1) / len(audit_inv_files), text=f"Processing {f.name}…")
                raw = f.read()

                inv_data   = extract_invoice_data(raw, f.name)
                invoice_no = inv_data.get("invoice_no")
                inv_date   = extract_invoice_date(raw)
                exp_due    = calc_due_date(inv_date)

                # Amounts from invoice PDF
                amts = extract_invoice_amounts(raw)
                inv_net   = float(amts["net"])   if amts["net"]   else None
                inv_taxes = amts["taxes"]  # already float or None
                inv_total = float(amts["total"]) if amts["total"] else None
                # If total not found, compute it
                if inv_total is None and inv_net is not None and inv_taxes is not None:
                    inv_total = round(inv_net + inv_taxes, 2)

                cc_prefix = inv_data.get("cc_prefix")
                if inv_data.get("is_six"):
                    vendor_ext = VENDOR_EXCEPCION
                    _, cc_ext  = get_vendor_cc(cc_prefix) if cc_prefix else (None, None)
                else:
                    vendor_ext, cc_ext = get_vendor_cc(cc_prefix) if cc_prefix else (None, None)
                gl_ext = get_gl(inv_data.get("product_code"))

                audit_rec    = audit_data.get(invoice_no) if invoice_no else None
                aud          = audit_rec or {}
                voucher_aud  = aud.get("voucher")
                inv_date_aud = aud.get("invoice_date")
                due_date_aud = aud.get("due_date")
                cc_aud       = aud.get("cc")
                vendor_aud   = aud.get("vendor")
                gl_aud       = aud.get("gl")
                aud_total    = aud.get("total")  # float or None (largest amount in audit row)

                total_ok = _cmp_amt(inv_total, aud_total)

                val_results.append({
                    "filename":     f.name,
                    "invoice_no":   invoice_no or "—",
                    "found":        audit_rec is not None,
                    # From invoice PDF
                    "inv_date":     inv_date,
                    "exp_due":      exp_due,
                    "cc_ext":       cc_ext,
                    "vendor_ext":   vendor_ext,
                    "gl_ext":       gl_ext,
                    "inv_net":      inv_net,
                    "inv_taxes":    inv_taxes,
                    "inv_total":    inv_total,
                    # From audit report
                    "voucher_aud":  voucher_aud,
                    "inv_date_aud": inv_date_aud,
                    "due_date_aud": due_date_aud,
                    "cc_aud":       cc_aud,
                    "vendor_aud":   vendor_aud,
                    "gl_aud":       gl_aud,
                    "aud_total":    aud_total,
                    # Validation flags
                    "inv_date_ok":  _cmp(inv_date, inv_date_aud),
                    "due_date_ok":  _cmp(exp_due, due_date_aud),
                    "cc_ok":        _cmp(cc_ext, cc_aud),
                    "vendor_ok":    _cmp(vendor_ext, vendor_aud),
                    "gl_ok":        _cmp(gl_ext, gl_aud),
                    "total_ok":     total_ok,
                })

            prog.progress(1.0, text="✅ Validation complete")
            st.session_state["audit_results"] = val_results
            st.rerun()

    # ── Results ────────────────────────────────────────────────────────────────
    if st.session_state.get("audit_results"):
        val_results  = st.session_state["audit_results"]
        audit_count  = st.session_state.get("audit_data_count", 0)

        def _icon(ok):
            if ok is None: return "❓"
            return "✅" if ok else "❌"

        def _fmt_date(d):
            return d.strftime("%d/%m/%Y") if d else "—"

        def _fmt_amt(v):
            if v is None: return "—"
            return f"{float(v):,.2f}"

        _all_flags = ["inv_date_ok", "due_date_ok", "cc_ok", "vendor_ok", "gl_ok", "total_ok"]

        n_found  = sum(1 for r in val_results if r["found"])
        n_all_ok = sum(
            1 for r in val_results
            if r["found"] and all(v is not False for v in [r[k] for k in _all_flags])
        )
        n_issues = len(val_results) - n_all_ok

        st.info(f"Audit report: **{audit_count}** record(s) parsed")
        st.markdown("<br>", unsafe_allow_html=True)

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(
                f"<div class='stat-box'><div class='stat-num blue'>{len(val_results)}</div>"
                f"<div class='stat-lbl'>Invoices checked</div></div>",
                unsafe_allow_html=True,
            )
        with c2:
            col_c = "green" if n_found == len(val_results) else "amber"
            st.markdown(
                f"<div class='stat-box'><div class='stat-num {col_c}'>{n_found}</div>"
                f"<div class='stat-lbl'>Found in audit</div></div>",
                unsafe_allow_html=True,
            )
        with c3:
            st.markdown(
                f"<div class='stat-box'><div class='stat-num green'>{n_all_ok}</div>"
                f"<div class='stat-lbl'>All fields OK</div></div>",
                unsafe_allow_html=True,
            )
        with c4:
            col_c = "red" if n_issues else "green"
            st.markdown(
                f"<div class='stat-box'><div class='stat-num {col_c}'>{n_issues}</div>"
                f"<div class='stat-lbl'>Need review</div></div>",
                unsafe_allow_html=True,
            )

        if n_issues == 0:
            st.success("✅ All invoices passed validation.")
        else:
            st.warning(f"⚠️ {n_issues} invoice(s) need review — see the Excel report below.")

        # ── Export to Excel ────────────────────────────────────────────────────
        st.divider()
        export_rows = []
        for r in val_results:
            export_rows.append({
                "File":                  r["filename"],
                "Invoice No":            r["invoice_no"],
                "Found in Audit":        "Yes" if r["found"] else "No",
                "Voucher No":            r.get("voucher_aud") or "",
                "Inv Date (PDF)":        _fmt_date(r["inv_date"]),
                "Inv Date (Audit)":      _fmt_date(r["inv_date_aud"]),
                "Inv Date OK":           _icon(r["inv_date_ok"]),
                "Payment Date (Calc)":   _fmt_date(r["exp_due"]),
                "Payment Date (Audit)":  _fmt_date(r["due_date_aud"]),
                "Payment Date OK":       _icon(r["due_date_ok"]),
                "CC (PDF)":              r["cc_ext"]    or "",
                "CC (Audit)":            r["cc_aud"]    or "",
                "CC OK":                 _icon(r["cc_ok"]),
                "Vendor (PDF)":          r["vendor_ext"] or "",
                "Vendor (Audit)":        r["vendor_aud"] or "",
                "Vendor OK":             _icon(r["vendor_ok"]),
                "GL (PDF)":              r["gl_ext"]    or "",
                "GL (Audit)":            r["gl_aud"]    or "",
                "GL OK":                 _icon(r["gl_ok"]),
                "Net/Subtotal (PDF)":    _fmt_amt(r["inv_net"]),
                "Taxes GST+QST (PDF)":  _fmt_amt(r["inv_taxes"]),
                "Total (PDF)":          _fmt_amt(r["inv_total"]),
                "Total (Audit)":        _fmt_amt(r["aud_total"]),
                "Total OK":             _icon(r["total_ok"]),
            })
        df_exp = pd.DataFrame(export_rows)
        buf_xl = BytesIO()
        with pd.ExcelWriter(buf_xl, engine="openpyxl") as xl_writer:
            df_exp.to_excel(xl_writer, index=False)
        buf_xl.seek(0)
        st.download_button(
            "⬇️ Download Validation Report (Excel)",
            data=buf_xl.read(),
            file_name=f"ap_audit_validation_{date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ══════════════════════════════════════════════════════════════════════════════
# TAB — STATEMENT RECONCILIATION
# ══════════════════════════════════════════════════════════════════════════════
if active_module == "recon":
    st.subheader("🧮 Statement Reconciliation")
    st.markdown(
        "Reconcile a vendor's **statement of account** against the accounting system "
        "and the invoice-copy tracking Excel. The tool walks every invoice on the "
        "statement and checks: is it registered in the system, does the amount match, "
        "do we have a copy on file, and does the PO on the statement match the PO on "
        "the file name."
    )

    vendor_keys = list(RECON_VENDORS.keys())
    recon_vendor = st.selectbox(
        "1️⃣ Vendor to reconcile",
        vendor_keys,
        format_func=lambda k: RECON_VENDORS[k]["label"],
    )
    st.caption(
        "The system extract and the invoice-copy tracking Excel are already scoped to "
        "this vendor when you export them — no extra filtering by name is applied; the "
        "vendor selection only decides which template is used to read the **statement**."
    )

    st.markdown("**2️⃣ Attach supporting files**")
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        recon_statement_file = st.file_uploader(
            "📄 Statement of Account (PDF or Excel)",
            type=["pdf", "xlsx"],
            key=f"recon_statement_{st.session_state.recon_upload_key}",
        )
    with col_b:
        recon_system_file = st.file_uploader(
            "🗄️ System extract (Excel)",
            type=["xlsx"],
            key=f"recon_system_{st.session_state.recon_upload_key}",
        )
    with col_c:
        recon_extraction_file = st.file_uploader(
            "📁 Invoice-copy tracking Excel (optional)",
            type=["xlsx"],
            key=f"recon_extraction_{st.session_state.recon_upload_key}",
        )

    col_run, col_clear = st.columns([1, 5])
    with col_run:
        do_recon = st.button(
            "🧮 Reconcile", type="primary",
            disabled=not (recon_statement_file and recon_system_file),
        )
    with col_clear:
        if st.button("🗑️ Clear", key="recon_clear"):
            st.session_state.recon_results = None
            st.session_state.recon_zip = None
            st.session_state.recon_upload_key += 1
            st.rerun()

    if do_recon:
        try:
            vendor_cfg = RECON_VENDORS[recon_vendor]
            check_po = vendor_cfg.get("check_po", True)
            statement, grand_total = vendor_cfg["parse_statement"](
                recon_statement_file.name, recon_statement_file.read()
            )
            system = parse_system_extract(recon_system_file.read())
            extraction = (
                parse_extraction_excel(recon_extraction_file.read())
                if recon_extraction_file else {}
            )
            if not statement:
                st.error("Could not find any invoice lines in the statement. Check the file format.")
            else:
                buckets = reconcile_statement(statement, system, extraction, check_po)
                statement_total = sum((r.get("amount") or 0) for r in statement)
                st.session_state.recon_results = buckets
                st.session_state.recon_statement_total = statement_total
                st.session_state.recon_grand_total = grand_total
                st.session_state.recon_check_po = check_po
                st.session_state.recon_zip = make_recon_report(buckets, statement_total, grand_total, check_po)
        except Exception as e:
            st.error(f"Error during reconciliation: {e}")

    buckets = st.session_state.recon_results
    if buckets:
        statement_total = st.session_state.recon_statement_total
        grand_total = st.session_state.recon_grand_total
        check_po = st.session_state.recon_check_po
        if grand_total is not None:
            diff = round(statement_total - grand_total, 2)
            if abs(diff) <= 0.02:
                st.success(
                    f"✅ Statement total matches: file says **CAD {grand_total:,.2f}**, "
                    f"parsed lines add up to **CAD {statement_total:,.2f}**."
                )
            else:
                st.error(
                    f"⚠️ Statement total does **not** match: file says **CAD {grand_total:,.2f}**, "
                    f"parsed lines add up to **CAD {statement_total:,.2f}** (difference "
                    f"**CAD {diff:,.2f}**) — the file may not have been read correctly."
                )
        else:
            st.warning(
                f"Could not find a total figure on the statement to check against — "
                f"parsed lines add up to **CAD {statement_total:,.2f}**."
            )

        st.markdown("<br>", unsafe_allow_html=True)
        summary_df = pd.DataFrame(
            [
                (label,
                 "" if count is None else f"{count:,}",
                 "" if amount is None else f"{amount:,.2f}")
                for label, count, amount in build_recon_summary_rows(buckets, statement_total, check_po)
            ],
            columns=["Category", "Invoice Count", "Amount (CAD)"],
        )
        st.dataframe(summary_df, use_container_width=True, hide_index=True)

        st.download_button(
            "⬇️ Download Reconciliation Report (Excel)",
            data=st.session_state.recon_zip,
            file_name=f"atlantic_reconciliation_{date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=False,
            type="primary",
        )
        st.caption(
            "Full invoice-level detail (including the **has copy** and **PO mismatch** "
            "columns on the Pending Registration sheet) is in the Excel report."
        )
    elif not (recon_statement_file or recon_system_file):
        st.info("📂 Upload the statement and the system extract to run the reconciliation.")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 6 — PAYMENT PACKAGER
# ══════════════════════════════════════════════════════════════════════════════
if active_module == "payment":
    st.subheader("💳 Payment Packager — Build the invoice packages for a payment run")
    st.markdown(
        "This tool works entirely through **upload / download** — it does **not** need access "
        "to your computer's folders or network drives, so it works from any browser without "
        "installing anything. Invoices are matched by **invoice number**: the first **8 "
        "characters** of each PDF's filename, against the **Reference** column (column **E**) "
        "of the AP payment report.\n\n"
        "1. Upload the **payment report** (Excel/CSV) — the list of invoices to pay.\n"
        "2. Add each **unpaid folder** separately (its name/vendor + all its current PDFs, "
        "paid or not) — one folder at a time, as many as you need.\n"
        "3. Click **Build payment packages** — you'll get a ZIP for the **Paid folder (AP / "
        "Vendors)**, a ZIP for the **payment-number folder (Finance)**, and — for each unpaid "
        "folder you added — a ready-to-swap-in **updated unpaid ZIP** with the paid invoices "
        "already removed, so you replace the whole folder instead of hunting files to delete."
    )

    def _pay_clean_ref(v) -> str:
        if v is None:
            return ""
        if isinstance(v, float) and v.is_integer():
            v = int(v)
        s = str(v).strip()
        if not s or s.lower() == "nan":
            return ""
        if re.fullmatch(r"\d+\.0+", s):
            s = s.split(".")[0]
        return s

    # ── 1. Payment report ────────────────────────────────────────────────────
    st.markdown("**1. Payment report (invoices to pay)**")
    pay_list_file = st.file_uploader(
        "Excel or CSV with the invoices selected for payment",
        type=["xlsx", "xls", "csv"],
        key="pay_list_upload",
    )

    pay_invoices = []
    if pay_list_file:
        df_pay = None
        try:
            if pay_list_file.name.lower().endswith(".csv"):
                df_pay = pd.read_csv(pay_list_file)
            else:
                df_pay = pd.read_excel(pay_list_file)
        except Exception as e:
            st.error(f"Could not read the file: {e}")

        if df_pay is not None and len(df_pay.columns):
            guess_idx = 4 if len(df_pay.columns) > 4 else 0
            for i, c in enumerate(df_pay.columns):
                if re.search(r"reference|referencia", str(c), re.I):
                    guess_idx = i
                    break
            pay_col = st.selectbox(
                "Column with the invoice number (defaults to column E — Reference)",
                options=list(df_pay.columns),
                index=guess_idx,
                key="pay_col_select",
            )
            st.caption(
                "Compared against the first 8 characters of each PDF's filename — "
                "e.g. `92006209_ML_24289954.pdf` → `92006209`."
            )
            pay_invoices = [
                _pay_clean_ref(v) for v in df_pay[pay_col].tolist()
            ]
            pay_invoices = [v for v in pay_invoices if v]
            st.caption(f"{len(pay_invoices)} invoice(s) loaded from **{pay_col}**")

    st.divider()

    # ── 2. Unpaid folders, added one at a time ──────────────────────────────
    st.markdown("**2. Unpaid folders**")
    st.caption(
        "Add every unpaid folder that might contain a selected invoice. Upload **all** the PDFs "
        "currently in that folder (not just the ones being paid) so the tool can hand back a "
        "complete, ready-to-swap-in replacement for it. If you upload everything in one single "
        "batch instead of folder by folder, any label works (e.g. `Unpaid`) — it's only used to "
        "name the ZIP you get back, it doesn't affect the matching."
    )

    existing_labels = {b["label"].lower() for b in st.session_state.payment_batches}
    _pay_form_key = st.session_state.payment_batch_form_key

    if st.session_state.payment_batches:
        _pay_total_mb = sum(
            len(f["bytes"]) for b in st.session_state.payment_batches for f in b["files"]
        ) / (1024 * 1024)
        cap_col, clr_col = st.columns([5, 1])
        with cap_col:
            st.caption(f"📦 {_pay_total_mb:,.1f} MB currently held in this session across all folders.")
        with clr_col:
            if st.button("🗑️ Clear all", key="pay_clear_all", use_container_width=True):
                st.session_state.payment_batches = []
                st.session_state.payment_result = None
                st.rerun()

    fc1, fc2 = st.columns([1, 2])
    with fc1:
        batch_label = st.text_input(
            "Folder name / vendor No.",
            placeholder="0101000430",
            key=f"pay_batch_label_{_pay_form_key}",
        )
    with fc2:
        batch_files = st.file_uploader(
            "All PDFs currently in that folder", type=["pdf"], accept_multiple_files=True,
            key=f"pay_batch_files_{_pay_form_key}",
        )
    add_batch = st.button("➕ Add this folder")

    if add_batch:
        label = batch_label.strip()
        if not label:
            st.warning("Enter a folder name / vendor number first.")
        elif label.lower() in existing_labels:
            st.warning(f"'{label}' was already added — remove it below first if you want to replace it.")
        elif not batch_files:
            st.warning("Upload that folder's PDFs first.")
        else:
            st.session_state.payment_batches.append({
                "label": label,
                "files": [{"name": f.name, "bytes": f.getvalue()} for f in batch_files],
            })
            st.session_state.payment_batch_form_key += 1
            st.rerun()

    if st.session_state.payment_batches:
        for i, b in enumerate(st.session_state.payment_batches):
            bc1, bc2 = st.columns([6, 1])
            with bc1:
                st.write(f"📁 **{b['label']}** — {len(b['files'])} file(s)")
            with bc2:
                if st.button("🗑️", key=f"pay_batch_del_{i}"):
                    st.session_state.payment_batches.pop(i)
                    st.rerun()
    else:
        st.info("No unpaid folders added yet.")

    st.divider()

    # ── 3. Payment number ─────────────────────────────────────────────────────
    st.markdown("**3. Payment number**")
    payment_no_txt = st.text_input(
        "Used to name the Finance package/folder",
        value="",
        placeholder="e.g. PAY-2026-08-17-001",
        key="pay_number_input",
    )

    st.divider()

    can_build = bool(pay_invoices and st.session_state.payment_batches and payment_no_txt.strip())
    if st.button("📦 Build payment packages", type="primary", disabled=not can_build):
        # Index every uploaded file across all folders, by invoice No. (first 8 chars of filename).
        upload_index = {}
        for b in st.session_state.payment_batches:
            for f in b["files"]:
                inv_key = f["name"][:8]
                upload_index.setdefault(inv_key, []).append(f)

        rows = []
        for inv_no in pay_invoices:
            matches = upload_index.get(inv_no, [])
            if len(matches) == 1:
                status = "found"
            elif len(matches) > 1:
                status = "duplicate"
            else:
                status = "not_found"
            rows.append({"invoice_no": inv_no, "matches": matches, "status": status})

        to_pack = [r for r in rows if r["status"] == "found"]
        paid_keys = {r["invoice_no"] for r in to_pack}

        # Paid and Finance folders get the exact same set of files — build the
        # ZIP once and reuse its bytes for both instead of duplicating it in memory.
        zip_paid_buf = BytesIO()
        with zipfile.ZipFile(zip_paid_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for r in to_pack:
                zf.writestr(r["matches"][0]["name"], r["matches"][0]["bytes"])
        zip_paid_bytes = zip_paid_buf.getvalue()

        payment_no_clean = payment_no_txt.strip()

        # Everything that will also go into the combined "download all" ZIP.
        all_entries = []
        for r in to_pack:
            all_entries.append((f"Facturas Pagadas (AP)/{r['matches'][0]['name']}", r["matches"][0]["bytes"]))
            all_entries.append((f"Finanzas - {payment_no_clean}/{r['matches'][0]['name']}", r["matches"][0]["bytes"]))

        # One "updated unpaid" ZIP per folder — everything except what just got paid.
        remaining_zips = []
        for b in st.session_state.payment_batches:
            remaining = [f for f in b["files"] if f["name"][:8] not in paid_keys]
            buf = BytesIO()
            with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
                for f in remaining:
                    zf.writestr(f["name"], f["bytes"])
            buf.seek(0)
            remaining_zips.append({
                "label": b["label"],
                "zip": buf.getvalue(),
                "removed": len(b["files"]) - len(remaining),
                "remaining": len(remaining),
            })
            safe_label = b["label"].replace("/", "-").replace("\\", "-")
            for f in remaining:
                all_entries.append((f"Unpaid - {safe_label}/{f['name']}", f["bytes"]))

        report_rows = [{
            "Invoice No":    r["invoice_no"],
            "File":          r["matches"][0]["name"] if len(r["matches"]) == 1 else
                              "; ".join(m["name"] for m in r["matches"]) if r["matches"] else "",
            "Status": {"found": "Packaged", "duplicate": "Duplicate — review",
                       "not_found": "Not found"}[r["status"]],
        } for r in rows]
        buf_report = BytesIO()
        with pd.ExcelWriter(buf_report, engine="openpyxl") as xl_writer:
            pd.DataFrame(report_rows).to_excel(xl_writer, index=False)
        buf_report.seek(0)
        all_entries.append((f"Checklist_{payment_no_clean}.xlsx", buf_report.getvalue()))

        zip_all_buf = BytesIO()
        with zipfile.ZipFile(zip_all_buf, "w", zipfile.ZIP_DEFLATED) as zf:
            for arcname, data in all_entries:
                zf.writestr(arcname, data)
        zip_all_buf.seek(0)

        st.session_state.payment_result = {
            "payment_no":     payment_no_clean,
            "rows":           [{
                "invoice_no": r["invoice_no"],
                "status":     r["status"],
                "files":      [m["name"] for m in r["matches"]],
            } for r in rows],
            "zip_paid":       zip_paid_bytes,
            "zip_finance":    zip_paid_bytes,
            "report_xlsx":    buf_report.getvalue(),
            "remaining_zips": remaining_zips,
            "zip_all":        zip_all_buf.getvalue(),
        }

    # ── Results ───────────────────────────────────────────────────────────────
    result = st.session_state.get("payment_result")
    if result:
        rows        = result["rows"]
        n_found     = sum(1 for r in rows if r["status"] == "found")
        n_dup       = sum(1 for r in rows if r["status"] == "duplicate")
        n_not_found = sum(1 for r in rows if r["status"] == "not_found")

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(
                f"<div class='stat-box'><div class='stat-num blue'>{len(rows)}</div>"
                f"<div class='stat-lbl'>Selected</div></div>", unsafe_allow_html=True)
        with c2:
            st.markdown(
                f"<div class='stat-box'><div class='stat-num green'>{n_found}</div>"
                f"<div class='stat-lbl'>Packaged</div></div>", unsafe_allow_html=True)
        with c3:
            col_c = "amber" if n_dup else "green"
            st.markdown(
                f"<div class='stat-box'><div class='stat-num {col_c}'>{n_dup}</div>"
                f"<div class='stat-lbl'>Duplicate — review</div></div>", unsafe_allow_html=True)
        with c4:
            col_c = "red" if n_not_found else "green"
            st.markdown(
                f"<div class='stat-box'><div class='stat-num {col_c}'>{n_not_found}</div>"
                f"<div class='stat-lbl'>Not found</div></div>", unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        if n_dup:
            st.warning(
                "⚠️ Some invoice numbers matched more than one uploaded PDF (same first 8 "
                "characters, possibly in different folders) — those are skipped automatically "
                "(in both the paid package and every 'updated unpaid' ZIP) so the wrong file "
                "isn't moved. Resolve them manually, then rebuild."
            )
        if n_not_found:
            st.info(
                "ℹ️ Invoice numbers marked **Not found** weren't among the uploaded PDFs — check "
                "the number or upload the right folder and build the packages again."
            )

        if n_found:
            st.success(f"✅ {n_found} invoice(s) packaged.")

            st.download_button(
                "⬇️ Download everything (ZIP)",
                data=result["zip_all"],
                file_name=f"pago_{result['payment_no']}_completo.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True,
            )
            st.caption(
                "Contains the Paid folder, the Finance folder, the checklist, and every "
                "updated unpaid folder — each in its own subfolder, ready to copy into place."
            )
            st.markdown("<br>", unsafe_allow_html=True)

            with st.expander("Or download each piece separately"):
                dcol1, dcol2, dcol3 = st.columns(3)
                with dcol1:
                    st.download_button(
                        "⬇️ Paid folder ZIP (AP / Vendors)",
                        data=result["zip_paid"],
                        file_name=f"facturas_pagadas_{date.today().strftime('%Y%m%d')}.zip",
                        mime="application/zip",
                        use_container_width=True,
                    )
                with dcol2:
                    st.download_button(
                        "⬇️ Finance ZIP (payment " + result["payment_no"] + ")",
                        data=result["zip_finance"],
                        file_name=f"pago_{result['payment_no']}.zip",
                        mime="application/zip",
                        use_container_width=True,
                    )
                with dcol3:
                    st.download_button(
                        "⬇️ Checklist (Excel)",
                        data=result["report_xlsx"],
                        file_name=f"payment_checklist_{result['payment_no']}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

                st.markdown("**📁 Updated unpaid folders — replace each folder's contents with its ZIP**")
                for rz in result["remaining_zips"]:
                    safe_label = re.sub(r"[^A-Za-z0-9_.-]+", "_", rz["label"])
                    st.download_button(
                        f"⬇️ Unpaid — {rz['label']}  ({rz['removed']} removed, {rz['remaining']} remaining)",
                        data=rz["zip"],
                        file_name=f"unpaid_{safe_label}_{date.today().strftime('%Y%m%d')}.zip",
                        mime="application/zip",
                        key=f"pay_remaining_dl_{safe_label}",
                        use_container_width=True,
                    )

            st.info(
                "💡 **Next steps:** unzip the Paid ZIP into your **AP / Vendors paid** folder, "
                "unzip the Finance ZIP into the **payment-number** folder, then for each unpaid "
                "folder listed above, **delete everything inside it and unzip its updated ZIP in "
                "its place** — no need to search for and delete individual files."
            )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 7 — DATABASE
# ══════════════════════════════════════════════════════════════════════════════
if active_module == "database":
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
            new_prov = []
            for r in edited_prov.to_dict("records"):
                pref = r.get("prefijo")
                if pref and str(pref).strip() and str(pref).strip().upper() != "NAN":
                    new_prov.append({
                        "prefijo": str(pref).strip().upper(),
                        "vendor":  str(r.get("vendor", "")).strip(),
                        "cc":      str(r.get("cc", "")).strip(),
                    })
            st.session_state.proveedores = new_prov
            st.success(f"✅ Vendor table updated — {len(new_prov)} vendors saved")
            st.rerun()

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
            new_gl = []
            for r in edited_gl.to_dict("records"):
                cod = r.get("codigo")
                gl  = r.get("gl")
                if cod and str(cod).strip() and str(cod).strip().upper() != "NAN":
                    new_gl.append({
                        "codigo": str(cod).strip().upper(),
                        "gl":     str(gl).strip() if gl and str(gl).strip().upper() != "NAN" else "",
                    })
            st.session_state.gl_codes = new_gl
            st.success(f"✅ GL table updated — {len(new_gl)} codes saved")
            st.rerun()

        st.divider()
        st.subheader("📥 Import from Excel (maestro_contable.xlsx)")
        xl_up = st.file_uploader("Upload Excel file", type=["xlsx"], key="xl_import")
        if xl_up:
            try:
                wb = openpyxl.load_workbook(xl_up)
                imported_vendors = 0
                imported_gl = 0

                if "proveedores" in wb.sheetnames:
                    ws = wb["proveedores"]
                    new_proveedores = []
                    for r in ws.iter_rows(min_row=2, values_only=True):
                        if len(r) >= 3 and r[0] not in (None, ""):
                            new_proveedores.append({
                                "prefijo": str(r[0]).strip(),
                                "vendor":  str(r[1]).strip() if r[1] is not None else "",
                                "cc":      str(r[2]).strip() if r[2] is not None else "",
                            })
                    if new_proveedores:
                        st.session_state.proveedores = new_proveedores
                        imported_vendors = len(new_proveedores)

                if "cuentas_gl" in wb.sheetnames:
                    ws = wb["cuentas_gl"]
                    new_gl = []
                    for r in ws.iter_rows(min_row=2, values_only=True):
                        if len(r) >= 2 and r[0] not in (None, "") and r[1] not in (None, ""):
                            new_gl.append({
                                "codigo": str(r[0]).strip(),
                                "gl":     str(r[1]).strip(),
                            })
                    if new_gl:
                        st.session_state.gl_codes = new_gl
                        imported_gl = len(new_gl)

                if imported_vendors == 0 and imported_gl == 0:
                    st.warning(
                        "⚠️ No data imported. Make sure the Excel has sheets named "
                        "**'proveedores'** and/or **'cuentas_gl'** with data starting on row 2."
                    )
                else:
                    st.success(
                        f"✅ Imported: {imported_vendors} vendors, {imported_gl} GL codes"
                    )
                    st.rerun()
            except Exception as e:
                st.error(f"Error importing Excel: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 8 — SETTINGS
# ══════════════════════════════════════════════════════════════════════════════
if active_module == "settings":
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
