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
import os
import shutil
from pathlib import Path
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
        # Payment Mover state
        "payment_unpaid_paths":  "",
        "payment_paid_folder":   "",
        "payment_finance_base":  "",
        "payment_recursive":     True,
        "payment_overwrite":     False,
        "payment_preview":       None,
        "payment_move_log":      None,
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
    Date search is limited to yk-5..yk+35 (APINV row + due-date row only).
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

    def _is_date_word(text: str) -> "date | None":
        t = text.strip()
        if _ISO_PAT.match(t):
            return _parse_date_str(t)
        for p in _OTHER_PATS:
            if p.match(t):
                return _parse_date_str(t)
        return None

    def _assign_dates(date_items: list):
        if not date_items:
            return None, None
        date_items.sort(key=lambda x: x[0])
        dates = [d for _, d in date_items]
        if len(dates) == 1:
            d = dates[0]
            return (None, d) if d.day in (15, 30) else (d, None)
        inv, due = dates[0], dates[1]
        if inv.day in (15, 30) and due.day not in (15, 30):
            inv, due = due, inv
        return inv, due

    try:
        # ── Merge all pages into one coordinate space ────────────────────────────
        all_words: list = []
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
                y_offset += page.height

        if not all_words:
            return records

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

            # Dates: APINV row ±15 above + 45 below.
            # Crystal Reports can place the date column up to ~12 pt above or
            # below the "APINV" text baseline; +45 captures the due-date line
            # even with generous line spacing.  ISO dates are specific enough
            # (YYYY-MM-DD) that false positives from sub-rows are not a risk.
            date_band_ys = [y for y in sorted_ys if yk - 15 <= y <= yk + 45]
            date_items: list = []
            for y in date_band_ys:
                for w in row_map.get(y, []):
                    d = _is_date_word(w["text"])
                    if d:
                        date_items.append((w["top"], d))
            inv_date, due_date = _assign_dates(date_items)

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
st.markdown("""
<h1 style='margin-bottom:0'>📄 Atlantic — Invoice Tools</h1>
<p style='color:gray;margin-top:4px'>Invoice Splitter &nbsp;·&nbsp; Invoice &amp; PO Matcher &nbsp;·&nbsp; Invoice Codifier</p>
""", unsafe_allow_html=True)

tab_split, tab_match, tab_cod, tab_couru, tab_audit, tab_pay, tab_db, tab_cfg = st.tabs([
    "✂️  Invoice Splitter",
    "🔗  Invoice Matcher",
    "🏷️  Invoice Coding",
    "📊  Couru Code",
    "🔍  AP Audit",
    "💳  Payment Mover",
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
# TAB 3 — INVOICE CODING (upload + review + results unified)
# ══════════════════════════════════════════════════════════════════════════════
with tab_cod:
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
with tab_couru:
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
with tab_audit:
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
# TAB 6 — PAYMENT MOVER
# ══════════════════════════════════════════════════════════════════════════════
with tab_pay:
    st.subheader("💳 Payment Mover — Move Selected Invoices to Payment Folders")
    st.markdown(
        "Upload the **Excel/CSV** listing the invoice filenames selected for this payment run. "
        "The tool searches one or more **unpaid** folders (recursively), then **copies each "
        "matched invoice** into the **AP / Vendors paid folder** and into a **Finance folder "
        "named after the payment number** — the original is removed from the unpaid folder only "
        "after both copies succeed."
    )
    st.warning(
        "⚠️ This only works when the app runs **locally** (or on a server) with direct access to "
        "the folder paths below — e.g. mapped network drives. It will **not** work on a "
        "browser-only / cloud-hosted deployment with no filesystem access."
    )

    # ── 1. Selected invoices file ────────────────────────────────────────────
    st.markdown("**1. Selected invoices file**")
    pay_list_file = st.file_uploader(
        "Excel or CSV with the filenames selected for payment",
        type=["xlsx", "xls", "csv"],
        key="pay_list_upload",
    )

    pay_filenames = []
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
            guess_idx = 0
            for i, c in enumerate(df_pay.columns):
                if re.search(r"archivo|file|factura|nombre|name", str(c), re.I):
                    guess_idx = i
                    break
            pay_col = st.selectbox(
                "Column with the invoice filename",
                options=list(df_pay.columns),
                index=guess_idx,
                key="pay_col_select",
            )
            st.caption("Values must match the file name exactly, including the extension (e.g. `.pdf`).")
            pay_filenames = [
                str(v).strip() for v in df_pay[pay_col].tolist()
                if v is not None and str(v).strip() and str(v).strip().lower() != "nan"
            ]
            st.caption(f"{len(pay_filenames)} filename(s) loaded from **{pay_col}**")

    st.divider()

    # ── 2. Source / destination folders ──────────────────────────────────────
    st.markdown("**2. Folders**")
    col_a, col_b = st.columns(2)
    with col_a:
        unpaid_paths_txt = st.text_area(
            "Unpaid folder(s) — one path per line",
            value=st.session_state.payment_unpaid_paths,
            height=100,
            placeholder="\\\\server\\share\\Unpaid\\Vendor0101000430\n\\\\server\\share\\Unpaid\\Vendor0101002430",
            key="pay_unpaid_paths_input",
        )
        pay_recursive = st.checkbox(
            "Search subfolders recursively",
            value=st.session_state.payment_recursive,
            key="pay_recursive_input",
        )
    with col_b:
        paid_folder_txt = st.text_input(
            "Paid invoices folder (AP / Vendors)",
            value=st.session_state.payment_paid_folder,
            key="pay_paid_folder_input",
        )
        finance_base_txt = st.text_input(
            "Finance base folder (a subfolder named with the payment number is created here)",
            value=st.session_state.payment_finance_base,
            key="pay_finance_base_input",
        )
        payment_no_txt = st.text_input(
            "Payment number",
            value="",
            placeholder="e.g. PAY-2026-08-17-001",
            key="pay_number_input",
        )
        pay_overwrite = st.checkbox(
            "Overwrite if the file already exists at destination",
            value=st.session_state.payment_overwrite,
            key="pay_overwrite_input",
        )

    # persist folder settings for convenience across reruns
    st.session_state.payment_unpaid_paths = unpaid_paths_txt
    st.session_state.payment_paid_folder  = paid_folder_txt
    st.session_state.payment_finance_base = finance_base_txt
    st.session_state.payment_recursive    = pay_recursive
    st.session_state.payment_overwrite    = pay_overwrite

    st.divider()

    # ── 3. Preview ────────────────────────────────────────────────────────────
    do_preview = st.button(
        "🔍 Preview matches",
        type="primary",
        disabled=not (pay_filenames and unpaid_paths_txt.strip()),
    )

    if do_preview:
        roots = [p.strip() for p in unpaid_paths_txt.splitlines() if p.strip()]
        index = {}  # lower(filename) -> [Path, ...]
        bad_roots = []
        for root in roots:
            root_path = Path(root)
            if not root_path.is_dir():
                bad_roots.append(root)
                continue
            if pay_recursive:
                walker = os.walk(root_path)
            else:
                walker = [(str(root_path), [], [f.name for f in root_path.iterdir() if f.is_file()])]
            for dirpath, _dirnames, filenames in walker:
                for fn in filenames:
                    index.setdefault(fn.lower(), []).append(Path(dirpath) / fn)

        if bad_roots:
            st.error(
                "The following folder(s) don't exist or aren't accessible:\n"
                + "\n".join(f"- {b}" for b in bad_roots)
            )

        preview_rows = []
        for name in pay_filenames:
            matches = index.get(name.lower(), [])
            if len(matches) == 1:
                status = "found"
            elif len(matches) > 1:
                status = "duplicate"
            else:
                status = "not_found"
            preview_rows.append({
                "filename": name,
                "matches": [str(m) for m in matches],
                "status": status,
            })
        st.session_state.payment_preview  = preview_rows
        st.session_state.payment_move_log = None

    # ── Preview results ───────────────────────────────────────────────────────
    if st.session_state.payment_preview:
        pay_rows    = st.session_state.payment_preview
        n_found     = sum(1 for r in pay_rows if r["status"] == "found")
        n_dup       = sum(1 for r in pay_rows if r["status"] == "duplicate")
        n_not_found = sum(1 for r in pay_rows if r["status"] == "not_found")

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.markdown(
                f"<div class='stat-box'><div class='stat-num blue'>{len(pay_rows)}</div>"
                f"<div class='stat-lbl'>Selected</div></div>", unsafe_allow_html=True)
        with c2:
            st.markdown(
                f"<div class='stat-box'><div class='stat-num green'>{n_found}</div>"
                f"<div class='stat-lbl'>Ready to move</div></div>", unsafe_allow_html=True)
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
        _pay_status_lbl = {"found": "✅ Ready", "duplicate": "⚠️ Duplicate", "not_found": "❌ Not found"}
        df_preview = pd.DataFrame([{
            "File":         r["filename"],
            "Status":       _pay_status_lbl[r["status"]],
            "Location(s)":  "\n".join(r["matches"]) if r["matches"] else "—",
        } for r in pay_rows])
        st.dataframe(df_preview, use_container_width=True, hide_index=True)

        if n_dup:
            st.warning(
                "⚠️ Duplicate filenames were found in more than one unpaid folder — these are "
                "skipped automatically. Resolve them manually, then re-run the preview."
            )
        if n_not_found:
            st.info("ℹ️ Files marked **Not found** are skipped — check the filename or the folder paths.")

        # ── 4. Move ───────────────────────────────────────────────────────────
        st.divider()
        st.markdown("**3. Move**")
        pay_confirm = st.checkbox(
            f"I confirm I want to move **{n_found}** invoice(s) out of the unpaid folder(s) into "
            f"the Paid and Finance folders.",
            key="pay_confirm_move",
        )
        pay_dest_ready = bool(paid_folder_txt.strip() and finance_base_txt.strip() and payment_no_txt.strip())
        can_move = pay_confirm and n_found > 0 and pay_dest_ready
        if n_found and not pay_dest_ready:
            st.info("ℹ️ Fill in the Paid folder, Finance base folder and Payment number to enable the move.")

        if st.button("📦 Move invoices now", type="primary", disabled=not can_move):
            paid_dir    = Path(paid_folder_txt.strip())
            finance_dir = Path(finance_base_txt.strip()) / payment_no_txt.strip()
            try:
                paid_dir.mkdir(parents=True, exist_ok=True)
                finance_dir.mkdir(parents=True, exist_ok=True)
            except Exception as e:
                st.error(f"Could not create destination folder(s): {e}")
                st.stop()

            move_log = []
            to_move  = [r for r in pay_rows if r["status"] == "found"]
            prog = st.progress(0, text="Moving…")
            for i, r in enumerate(to_move):
                prog.progress(i / max(len(to_move), 1), text=f"Moving {r['filename']}…")
                src          = Path(r["matches"][0])
                dest_paid    = paid_dir / src.name
                dest_finance = finance_dir / src.name
                entry = {
                    "filename":     src.name,
                    "source":       str(src),
                    "paid_dest":    str(dest_paid),
                    "finance_dest": str(dest_finance),
                    "status":       "",
                    "detail":       "",
                }
                try:
                    if not pay_overwrite and (dest_paid.exists() or dest_finance.exists()):
                        entry["status"] = "skipped"
                        entry["detail"] = "Already exists at destination"
                    else:
                        shutil.copy2(src, dest_paid)
                        shutil.copy2(src, dest_finance)
                        os.remove(src)
                        entry["status"] = "moved"
                except Exception as e:
                    entry["status"] = "error"
                    entry["detail"] = str(e)
                move_log.append(entry)
            prog.progress(1.0, text="✅ Done")
            st.session_state.payment_move_log = move_log
            st.rerun()

    # ── Move results / log ───────────────────────────────────────────────────
    if st.session_state.payment_move_log:
        move_log  = st.session_state.payment_move_log
        n_moved   = sum(1 for e in move_log if e["status"] == "moved")
        n_skipped = sum(1 for e in move_log if e["status"] == "skipped")
        n_error   = sum(1 for e in move_log if e["status"] == "error")

        if n_error:
            st.error(f"⚠️ {n_moved} moved, {n_skipped} skipped, {n_error} error(s) — see the log below.")
        elif n_skipped:
            st.warning(f"✅ {n_moved} moved, {n_skipped} skipped (already existed at destination).")
        else:
            st.success(f"✅ {n_moved} invoice(s) moved to the Paid and Finance folders.")

        df_log = pd.DataFrame(move_log)
        st.dataframe(df_log, use_container_width=True, hide_index=True)

        buf_log = BytesIO()
        with pd.ExcelWriter(buf_log, engine="openpyxl") as xl_writer:
            df_log.to_excel(xl_writer, index=False)
        buf_log.seek(0)
        st.download_button(
            "⬇️ Download move log (Excel)",
            data=buf_log.read(),
            file_name=f"payment_move_log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ══════════════════════════════════════════════════════════════════════════════
# TAB 7 — DATABASE
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
