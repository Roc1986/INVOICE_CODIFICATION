"""
Codificador de Facturas — v1.1
Automatización de codificación GL/CC para facturas PDF (Atlantic Packaging)
Soporte para PDFs digitales y PDFs escaneados (OCR)
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

# OCR para PDFs escaneados (importación opcional con fallback)
try:
    import pytesseract
    from pdf2image import convert_from_bytes
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Codificador de Facturas",
    layout="wide",
    page_icon="📄",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# DATOS POR DEFECTO  (extraídos del maestro_contable.xlsx)
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
VENDOR_EXCEPCION = "0101000390"  # Vendor fijo para facturas que comienzan con "6"

# ─────────────────────────────────────────────────────────────────────────────
# INICIALIZACIÓN DE SESSION STATE
# ─────────────────────────────────────────────────────────────────────────────
def init_state():
    defaults = {
        "proveedores":  copy.deepcopy(DEFAULT_PROVEEDORES),
        "gl_codes":     copy.deepcopy(DEFAULT_GL_CODES),
        "usuarios":     DEFAULT_USERS.copy(),
        "processed":    [],   # [{filename, original_bytes, pdf_bytes, meta…}]
        "stamp_x":      281,   # centrado: (792-230)/2 — landscape 792x612
        "stamp_y_top":  594,   # casi en el tope: 612-18
        "stamp_w":      230,
        "stamp_h":      82,
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

init_state()

# ─────────────────────────────────────────────────────────────────────────────
# FUNCIONES CORE
# ─────────────────────────────────────────────────────────────────────────────

def get_vendor_cc(prefix: str):
    """Busca vendor y CC dado un prefijo de 2 letras."""
    prefix = prefix.upper()
    for row in st.session_state.proveedores:
        if str(row["prefijo"]).upper() == prefix:
            return str(row["vendor"]), str(row["cc"])
    return None, None


def get_gl(product_code: str) -> str | None:
    """Busca GL dado un código de producto."""
    if not product_code:
        return None
    code = product_code.upper().strip()
    for row in st.session_state.gl_codes:
        if str(row["codigo"]).upper().strip() == code:
            return str(row["gl"])
    return None


def extract_invoice_data(pdf_bytes: bytes, filename: str = "") -> dict:
    """
    Extrae del PDF (página 1):
      - invoice_no   : número de factura (ej. "82196527")
      - customer_order: N° orden cliente (ej. "ML11590")
      - cc_prefix    : prefijo de 2 letras (ej. "ML")
      - product_code : código de producto (ej. "LB2035")
      - is_six       : True si el N° de factura empieza con "6"
    Soporta PDFs digitales y PDFs escaneados (con OCR automático).
    """
    result = {
        "invoice_no": None,
        "customer_order": None,
        "cc_prefix": None,
        "product_code": None,
        "is_six": False,
        "raw_lines": [],
        "ocr_used": False,
        "error": None,
    }
    try:
        # ── Paso 1: intentar extracción de texto digital ──────────────────
        with pdfplumber.open(BytesIO(pdf_bytes)) as pdf:
            if not pdf.pages:
                result["error"] = "PDF vacío"
                return result
            text = pdf.pages[0].extract_text() or ""

        # ── Paso 2: si no hay texto, usar OCR ────────────────────────────
        if len(text.strip()) < 50:
            if OCR_AVAILABLE:
                try:
                    images = convert_from_bytes(pdf_bytes, first_page=1, last_page=1, dpi=200)
                    text = pytesseract.image_to_string(images[0])
                    result["ocr_used"] = True
                except Exception as ocr_err:
                    result["error"] = f"OCR falló: {ocr_err}"
                    return result
            else:
                result["error"] = "PDF escaneado — OCR no disponible en este entorno"
                return result

        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        result["raw_lines"] = lines
        text_upper = text.upper()

        # ── 3. Número de factura ──────────────────────────────────────────
        # Formato A (digital): número en línea ANTES de "INVOICE No/No DE FACTURE"
        # Formato B (OCR/portrait): número en la MISMA línea → "INVOICE No/No DE FACTURE 87060660"
        for i, line in enumerate(lines):
            if "INVOICE NO" in line.upper() and "FACTURE" in line.upper():
                # Formato B: número al final de la misma línea
                m = re.search(r"(\d{7,10})\s*$", line)
                if m:
                    result["invoice_no"] = m.group(1)
                    result["is_six"] = result["invoice_no"].startswith("6")
                    break
                # Formato A: número en líneas anteriores
                for j in range(max(0, i - 5), i):
                    if re.fullmatch(r"\d{6,10}", lines[j]):
                        result["invoice_no"] = lines[j]
                        result["is_six"] = result["invoice_no"].startswith("6")
                        break
                break

        # Fallback: primer número en el nombre del archivo ("61037398_EV_...pdf")
        if not result["invoice_no"] and filename:
            m = re.match(r"(\d{6,10})", filename)
            if m:
                result["invoice_no"] = m.group(1)
                result["is_six"] = result["invoice_no"].startswith("6")

        # ── 4. Customer Order No ──────────────────────────────────────────
        # Línea: "635108 100 ML11694 ..."  o  "635108 100 | ML11694 ..." (OCR)
        for line in lines:
            m = re.search(r"\d{6}\s+\d{2,3}\s+\|?\s*([A-Z]{2}\d{4,7})\b", line)
            if m:
                result["customer_order"] = m.group(1)
                result["cc_prefix"] = m.group(1)[:2].upper()
                break

        # ── 5. Código de producto ─────────────────────────────────────────
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


def create_stamp(user: str, vendor: str, cc: str, gl: str,
                 coding_date, page_w: float, page_h: float,
                 rotation: int = 0) -> bytes:
    """
    Genera el sello rojo centrado en la parte superior de la página,
    teniendo en cuenta la rotación del PDF (/Rotate).
    """
    sw = st.session_state.stamp_w
    sh = st.session_state.stamp_h
    margin = 18  # margen desde el borde superior visible

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
        # Página rotada: el contenido se muestra landscape (ancho=page_h, alto=page_w)
        disp_w = page_h   # ancho visible
        disp_h = page_w   # alto visible

        # Centro del sello en coordenadas de pantalla
        disp_cx = disp_w / 2                    # centro horizontal
        disp_cy = disp_h - margin - sh / 2      # cerca del borde superior

        if rotation == 90:
            # /Rotate=90 (CCW): x_pdf = pw - y_disp, y_pdf = x_disp
            cx_pdf = page_w - disp_cy
            cy_pdf = disp_cx
            rot_angle = 90    # CCW para compensar el /Rotate=90 CW de la página
        else:  # 270
            # /Rotate=270 (CW): x_pdf = y_disp, y_pdf = pw - x_disp
            cx_pdf = disp_cy
            cy_pdf = page_w - disp_cx
            rot_angle = 90

        c.saveState()
        c.translate(cx_pdf, cy_pdf)
        c.rotate(rot_angle)

        # Dibujar sello centrado en el origen local
        lx = -sw / 2
        ly = -sh / 2
        c.setStrokeColorRGB(0.85, 0.0, 0.0)
        c.setFillColorRGB(1.0, 1.0, 1.0)
        c.setLineWidth(1.8)
        c.rect(lx, ly, sw, sh, fill=1)
        c.setFillColorRGB(0.85, 0.0, 0.0)
        c.setFont("Helvetica-Bold", 8.5)
        line_h = sh / 5.2
        tx = lx + 10
        ty = ly + sh - line_h
        for i, line in enumerate(stamp_lines):
            c.drawString(tx, ty - i * line_h, line)
        c.restoreState()

    else:
        # Página normal (sin rotación o 180°): sello centrado arriba
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
        tx = sx + 10
        ty = sy_bot + sh - line_h
        for i, line in enumerate(stamp_lines):
            c.drawString(tx, ty - i * line_h, line)

    c.save()
    packet.seek(0)
    return packet.read()


def stamp_pdf(original_bytes: bytes, stamp_bytes: bytes) -> bytes:
    """Superpone el sello sobre la página 1 del PDF original."""
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


def process_one(original_bytes: bytes, user: str, vendor: str,
                cc: str, gl: str, coding_date) -> bytes:
    """Lee las dimensiones y rotación de la página, crea el sello y lo aplica."""
    reader = PdfReader(BytesIO(original_bytes))
    page = reader.pages[0]
    pw = float(page.mediabox.width)
    ph = float(page.mediabox.height)
    rotation = int(page.get("/Rotate", 0) or 0)
    stamp_bytes = create_stamp(user, vendor, cc, gl, coding_date, pw, ph, rotation)
    return stamp_pdf(original_bytes, stamp_bytes)


def make_zip(items: list) -> bytes:
    """Genera un ZIP con todos los PDFs procesados (mismo nombre que el original)."""
    buf = BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for item in items:
            zf.writestr(item["filename"], item["pdf_bytes"])  # nombre original sin prefijo
    buf.seek(0)
    return buf.read()

# ─────────────────────────────────────────────────────────────────────────────
# ESTILOS CSS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* Tabs más grandes */
.stTabs [data-baseweb="tab"] { font-size: 15px; font-weight: 600; padding: 8px 18px; }

/* Tarjetas de factura */
.inv-card {
    background: #f8f9fa; border-radius: 10px; padding: 14px 18px;
    margin-bottom: 8px; border-left: 5px solid #ccc;
}
.inv-ok   { border-left-color: #28a745 !important; }
.inv-warn { border-left-color: #ffc107 !important; }
.inv-err  { border-left-color: #dc3545 !important; }

/* Sello preview */
.stamp-preview {
    border: 2px solid #cc0000; padding: 8px 12px;
    display: inline-block; background: white;
    font-family: monospace; font-weight: bold; color: #cc0000;
    font-size: 13px; line-height: 1.6; border-radius: 3px;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## ⚙️ Sesión de Trabajo")

    # Selección de usuario
    user_list = st.session_state.usuarios + ["✏️ Otro..."]
    sel_idx = st.selectbox("👤 Responsable (Posted By)", range(len(user_list)),
                           format_func=lambda x: user_list[x])
    if user_list[sel_idx] == "✏️ Otro...":
        current_user = st.text_input("Nombre:", placeholder="Escribe el nombre")
    else:
        current_user = user_list[sel_idx]

    # Fecha
    coding_date = st.date_input("📅 Fecha de codificación", value=date.today())

    st.divider()

    # Métricas rápidas
    n_proc = len(st.session_state.processed)
    st.metric("Facturas codificadas", n_proc)

    if n_proc > 0:
        zip_bytes = make_zip(st.session_state.processed)
        st.download_button(
            "⬇️ Descargar ZIP (todas)",
            data=zip_bytes,
            file_name=f"facturas_{date.today().strftime('%Y%m%d')}.zip",
            mime="application/zip",
            use_container_width=True,
        )
        if st.button("🗑️ Borrar todos los resultados", use_container_width=True, type="secondary"):
            st.session_state.processed = []
            st.rerun()

    st.divider()

    # Import/Export BD
    with st.expander("💾 Guardar / Cargar Base de Datos"):
        db_export = {
            "proveedores": st.session_state.proveedores,
            "gl_codes":    st.session_state.gl_codes,
            "usuarios":    st.session_state.usuarios,
        }
        st.download_button(
            "⬇️ Exportar BD (JSON)",
            data=json.dumps(db_export, indent=2, ensure_ascii=False),
            file_name="bd_facturas.json",
            mime="application/json",
            use_container_width=True,
        )
        db_upload = st.file_uploader("📥 Importar BD (JSON)", type=["json"], key="db_import")
        if db_upload:
            try:
                db = json.loads(db_upload.read())
                if "proveedores" in db: st.session_state.proveedores = db["proveedores"]
                if "gl_codes"    in db: st.session_state.gl_codes    = db["gl_codes"]
                if "usuarios"    in db: st.session_state.usuarios     = db["usuarios"]
                st.success("✅ Base de datos cargada")
            except Exception as e:
                st.error(f"Error: {e}")

# ─────────────────────────────────────────────────────────────────────────────
# CABECERA PRINCIPAL
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<h1 style='margin-bottom:0'>📄 Codificador de Facturas</h1>
<p style='color:gray;margin-top:4px'>Codificación automática GL / Centro de Costo en facturas PDF</p>
""", unsafe_allow_html=True)

tab_proc, tab_res, tab_db, tab_cfg = st.tabs([
    "📤  Procesar Facturas",
    "📋  Resultados",
    "🗄️  Base de Datos",
    "⚙️  Configuración",
])

# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — PROCESAR FACTURAS
# ══════════════════════════════════════════════════════════════════════════════
with tab_proc:
    st.subheader("Subir Facturas PDF")

    # Clave dinámica para poder limpiar el uploader
    if "upload_key" not in st.session_state:
        st.session_state.upload_key = 0

    col_up, col_clear = st.columns([5, 1])
    with col_up:
        uploaded = st.file_uploader(
            "Arrastra o selecciona uno o más PDFs",
            type=["pdf"],
            accept_multiple_files=True,
            help="Soporta carga masiva. Se procesa solo la primera página de cada factura.",
            key=f"uploader_{st.session_state.upload_key}",
        )
    with col_clear:
        st.markdown("<br>", unsafe_allow_html=True)  # alinea verticalmente
        if st.button("🗑️ Limpiar\ncarga", use_container_width=True,
                     help="Elimina todos los archivos cargados para subir un nuevo lote"):
            st.session_state.upload_key += 1
            st.rerun()

    if not uploaded:
        st.info("📂 Sube facturas para comenzar. Puedes seleccionar múltiples archivos a la vez.")
    else:
        st.success(f"✅ **{len(uploaded)} archivo(s)** cargado(s) — revisando…")
        st.divider()

        # ── Análisis de cada factura ──────────────────────────────────────
        invoices_ui = []   # lista de dicts con datos + widgets state keys

        for idx, f in enumerate(uploaded):
            raw = f.read()
            data = extract_invoice_data(raw, f.name)
            data["filename"] = f.name
            data["raw_bytes"] = raw

            # Determinar CC, Vendor, GL según lógica de negocio
            inv_no = data.get("invoice_no") or ""
            is_six = data.get("is_six", False)

            if is_six:
                data["vendor_auto"] = VENDOR_EXCEPCION
                data["cc_auto"]     = None
                data["needs_cc"]    = True
            else:
                prefix = data.get("cc_prefix")
                v, c = get_vendor_cc(prefix) if prefix else (None, None)
                data["vendor_auto"] = v
                data["cc_auto"]     = c
                data["needs_cc"]    = (v is None or c is None)

            data["gl_auto"] = get_gl(data.get("product_code"))
            invoices_ui.append(data)

        # ── Mostrar tarjetas de revisión ──────────────────────────────────
        st.subheader("Revisión — Datos extraídos")

        resolved_cc     = {}   # idx → cc final
        resolved_vendor = {}   # idx → vendor final
        resolved_gl     = {}   # idx → gl final

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

                # ── Col 1: datos extraídos
                with c1:
                    st.markdown("**📑 Extraído del PDF**")
                    st.write(f"N° Factura: `{inv.get('invoice_no') or '—'}`")
                    st.write(f"Orden cliente: `{inv.get('customer_order') or '—'}`")
                    st.write(f"Prefijo CC: `{inv.get('cc_prefix') or '—'}`")
                    st.write(f"Código producto: `{inv.get('product_code') or '—'}`")
                    if inv.get("is_six"):
                        st.info("⚡ Factura tipo **6** — excepción de vendor")
                    if inv.get("error"):
                        st.error(f"Error: {inv['error']}")
                    if inv.get("ocr_used"):
                        st.info("🔍 Texto extraído con OCR (PDF escaneado)")

                # ── Col 2: CC y Vendor
                with c2:
                    st.markdown("**🏷️ Vendor / Centro de Costo**")

                    if inv["is_six"]:
                        st.write(f"Vendor (fijo): `{VENDOR_EXCEPCION}`")
                        resolved_vendor[idx] = VENDOR_EXCEPCION
                        # CC manual para facturas tipo 6
                        cc_opts = sorted(set(r["cc"] for r in st.session_state.proveedores))
                        sel_cc = st.selectbox(
                            "Selecciona CC:", cc_opts, key=f"cc6_{idx}"
                        )
                        resolved_cc[idx] = sel_cc
                        st.success(f"CC seleccionado: `{sel_cc}`")

                    elif inv["vendor_auto"] and inv["cc_auto"]:
                        st.success(f"Vendor: `{inv['vendor_auto']}`")
                        st.success(f"CC: `{inv['cc_auto']}`")
                        resolved_vendor[idx] = inv["vendor_auto"]
                        resolved_cc[idx]     = inv["cc_auto"]

                    else:
                        st.warning(f"Prefijo `{inv.get('cc_prefix')}` no encontrado")
                        all_opts = sorted(set(r["cc"] for r in st.session_state.proveedores))
                        sel_cc = st.selectbox("CC manual:", all_opts, key=f"ccman_{idx}")
                        sel_vendor = next(
                            (r["vendor"] for r in st.session_state.proveedores if r["cc"] == sel_cc),
                            "DESCONOCIDO"
                        )
                        resolved_cc[idx]     = sel_cc
                        resolved_vendor[idx] = sel_vendor

                # ── Col 3: GL
                with c3:
                    st.markdown("**📊 Cuenta GL**")
                    if inv.get("gl_auto"):
                        st.success(f"GL: `{inv['gl_auto']}`")
                        resolved_gl[idx] = inv["gl_auto"]
                    else:
                        st.warning("GL no detectado automáticamente")
                        gl_opts = sorted(set(r["gl"] for r in st.session_state.gl_codes))
                        sel_gl = st.selectbox("GL manual:", gl_opts, key=f"glman_{idx}")
                        resolved_gl[idx] = sel_gl

                # ── Preview del sello
                cc_prev = resolved_cc.get(idx, "??")
                gl_prev = resolved_gl.get(idx, "??")
                vd_prev = resolved_vendor.get(idx, "??")
                usr_prev = current_user or "???"
                date_prev = coding_date.strftime("%d/%m/%Y")
                st.markdown(f"""
                <div style='margin-top:10px'>
                <p style='margin-bottom:4px; color:gray; font-size:12px'>👁️ Preview del sello:</p>
                <div class='stamp-preview'>
                POSTED BY: {usr_prev}<br>
                VENDOR: {vd_prev}<br>
                CC: {cc_prev}&nbsp;&nbsp;|&nbsp;&nbsp;GL: {gl_prev}<br>
                DATE: {date_prev}
                </div></div>
                """, unsafe_allow_html=True)

        # ── Botón Procesar ─────────────────────────────────────────────────
        st.divider()
        col_btn, col_info = st.columns([1, 3])
        with col_btn:
            do_process = st.button(
                "🚀 Procesar Facturas",
                type="primary",
                use_container_width=True,
                disabled=not bool(current_user),
            )
        with col_info:
            if not current_user:
                st.warning("⚠️ Selecciona un responsable (Posted By) en la barra lateral antes de procesar.")

        if do_process and current_user:
            progress = st.progress(0, text="Iniciando…")
            errors = []

            for idx, inv in enumerate(invoices_ui):
                fname = inv["filename"]
                progress.progress((idx + 1) / len(invoices_ui), text=f"Procesando {fname}…")

                cc     = resolved_cc.get(idx, "???")
                vendor = resolved_vendor.get(idx, "???")
                gl     = resolved_gl.get(idx, "???")

                try:
                    stamped = process_one(
                        inv["raw_bytes"], current_user,
                        vendor, cc, gl, coding_date,
                    )
                    st.session_state.processed.append({
                        "filename":      fname,
                        "original_bytes": inv["raw_bytes"],
                        "pdf_bytes":     stamped,
                        "invoice_no":    inv.get("invoice_no"),
                        "vendor":        vendor,
                        "cc":            cc,
                        "gl":            gl,
                        "user":          current_user,
                        "date":          coding_date.strftime("%d/%m/%Y"),
                        "date_obj":      coding_date,
                        "ts":            datetime.now().strftime("%Y-%m-%d %H:%M"),
                    })
                except Exception as e:
                    errors.append(f"{fname}: {e}")

            progress.progress(1.0, text="✅ Completado")

            if errors:
                for err in errors:
                    st.error(err)
            else:
                st.success(f"🎉 **{len(invoices_ui)} factura(s)** codificada(s) exitosamente. Ve a la pestaña **Resultados** para descargarlas.")
                st.balloons()

# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — RESULTADOS
# ══════════════════════════════════════════════════════════════════════════════
with tab_res:
    if not st.session_state.processed:
        st.info("📭 No hay facturas procesadas todavía. Ve a **Procesar Facturas** para comenzar.")
    else:
        n = len(st.session_state.processed)
        st.subheader(f"Facturas Procesadas — {n} archivo(s)")

        # ── Descarga masiva ───────────────────────────────────────────────
        col_dl, col_del, _ = st.columns([2, 2, 5])
        with col_dl:
            zip_all = make_zip(st.session_state.processed)
            st.download_button(
                f"⬇️ Descargar ZIP ({n} facturas)",
                data=zip_all,
                file_name=f"facturas_codificadas_{date.today().strftime('%Y%m%d')}.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True,
            )
        with col_del:
            if st.button("🗑️ Borrar todos", use_container_width=True):
                st.session_state.processed = []
                st.rerun()

        st.divider()

        # ── Modificar fecha ───────────────────────────────────────────────
        with st.expander("📅 Modificar Fecha de Codificación"):
            st.caption("Usa esto si la factura no pudo registrarse el mismo día que se codeó.")
            col_nd, col_sel = st.columns([1, 2])
            with col_nd:
                new_date = st.date_input("Nueva fecha:", value=date.today(), key="new_date_picker")
            with col_sel:
                inv_labels = [
                    f"{item['filename']}  ({item['date']})"
                    for item in st.session_state.processed
                ]
                sel_labels = st.multiselect("Selecciona facturas a actualizar:", inv_labels)

            if st.button("🔄 Regenerar con nueva fecha", type="primary"):
                count = 0
                for i, item in enumerate(st.session_state.processed):
                    label = f"{item['filename']}  ({item['date']})"
                    if label in sel_labels:
                        try:
                            new_stamped = process_one(
                                item["original_bytes"],
                                item["user"], item["vendor"],
                                item["cc"], item["gl"], new_date,
                            )
                            st.session_state.processed[i]["pdf_bytes"] = new_stamped
                            st.session_state.processed[i]["date"] = new_date.strftime("%d/%m/%Y")
                            st.session_state.processed[i]["date_obj"] = new_date
                            count += 1
                        except Exception as e:
                            st.error(f"Error en {item['filename']}: {e}")
                if count:
                    st.success(f"✅ {count} factura(s) regenerada(s) con fecha {new_date.strftime('%d/%m/%Y')}")
                    st.rerun()

        st.divider()

        # ── Lista de facturas procesadas ──────────────────────────────────
        to_delete = []
        for i, item in enumerate(st.session_state.processed):
            col1, col2, col3, col4 = st.columns([4, 1.5, 1, 1])
            with col1:
                st.markdown(f"📄 **{item['filename']}**")
                st.caption(
                    f"Vendor: `{item['vendor']}` | CC: `{item['cc']}` | "
                    f"GL: `{item['gl']}` | By: **{item['user']}** | "
                    f"Fecha: `{item['date']}` | Procesado: {item['ts']}"
                )
            with col2:
                st.download_button(
                    "⬇️ Descargar",
                    data=item["pdf_bytes"],
                    file_name=item['filename'],
                    mime="application/pdf",
                    key=f"dl_{i}",
                    use_container_width=True,
                )
            with col3:
                if st.button("🗑️", key=f"del_{i}", help="Eliminar de la lista"):
                    to_delete.append(i)
            with col4:
                st.caption(f"#{i+1}")
            st.markdown("<hr style='margin:6px 0; border-color:#eee'>", unsafe_allow_html=True)

        # Aplicar eliminaciones
        if to_delete:
            for i in sorted(to_delete, reverse=True):
                st.session_state.processed.pop(i)
            st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — BASE DE DATOS
# ══════════════════════════════════════════════════════════════════════════════
with tab_db:
    db_t1, db_t2 = st.tabs(["🏢 Proveedores / CC", "📊 Cuentas GL"])

    # ── Proveedores ────────────────────────────────────────────────────────
    with db_t1:
        st.subheader("Tabla de Proveedores")
        st.caption("Relaciona el **prefijo** (2 letras del N° de Orden) → Vendor y Centro de Costo")

        df_prov = pd.DataFrame(st.session_state.proveedores)
        edited_prov = st.data_editor(
            df_prov,
            column_config={
                "prefijo": st.column_config.TextColumn("Prefijo", width="small",
                    help="2 letras iniciales del N° de Orden del cliente"),
                "vendor":  st.column_config.TextColumn("N° Vendor", width="medium"),
                "cc":      st.column_config.TextColumn("Centro de Costo", width="medium"),
            },
            num_rows="dynamic",
            use_container_width=True,
            key="prov_editor",
        )
        if st.button("💾 Guardar cambios — Proveedores", type="primary"):
            st.session_state.proveedores = [
                r for r in edited_prov.to_dict("records")
                if r.get("prefijo")
            ]
            st.success("✅ Tabla de proveedores actualizada")

    # ── Cuentas GL ────────────────────────────────────────────────────────
    with db_t2:
        st.subheader("Tabla de Códigos GL")
        st.caption("Relaciona el **código de producto** de la factura → Cuenta GL")

        df_gl = pd.DataFrame(st.session_state.gl_codes)
        edited_gl = st.data_editor(
            df_gl,
            column_config={
                "codigo": st.column_config.TextColumn("Código de Producto", width="medium"),
                "gl":     st.column_config.TextColumn("Cuenta GL", width="medium"),
            },
            num_rows="dynamic",
            use_container_width=True,
            key="gl_editor",
        )
        if st.button("💾 Guardar cambios — GL", type="primary"):
            st.session_state.gl_codes = [
                r for r in edited_gl.to_dict("records")
                if r.get("codigo")
            ]
            st.success("✅ Tabla GL actualizada")

        # Importar desde Excel
        st.divider()
        st.subheader("📥 Importar desde Excel (maestro_contable.xlsx)")
        xl_up = st.file_uploader("Sube el archivo Excel", type=["xlsx"], key="xl_import")
        if xl_up:
            try:
                wb = openpyxl.load_workbook(xl_up)

                if "proveedores" in wb.sheetnames:
                    ws = wb["proveedores"]
                    rows = list(ws.iter_rows(min_row=2, values_only=True))
                    new_p = [
                        {"prefijo": str(r[0]), "vendor": str(r[1]), "cc": str(r[2])}
                        for r in rows if r[0]
                    ]
                    st.session_state.proveedores = new_p

                if "cuentas_gl" in wb.sheetnames:
                    ws = wb["cuentas_gl"]
                    rows = list(ws.iter_rows(min_row=2, values_only=True))
                    new_g = [
                        {"codigo": str(r[0]), "gl": str(r[1])}
                        for r in rows if r[0] and r[1]
                    ]
                    st.session_state.gl_codes = new_g

                st.success(
                    f"✅ Importado: {len(st.session_state.proveedores)} proveedores, "
                    f"{len(st.session_state.gl_codes)} códigos GL"
                )
                st.rerun()
            except Exception as e:
                st.error(f"Error al importar Excel: {e}")

# ══════════════════════════════════════════════════════════════════════════════
# TAB 4 — CONFIGURACIÓN
# ══════════════════════════════════════════════════════════════════════════════
with tab_cfg:
    st.subheader("Configuración General")
    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**👥 Usuarios del sistema**")
        users_txt = st.text_area(
            "Un usuario por línea:",
            value="\n".join(st.session_state.usuarios),
            height=130,
        )
        if st.button("💾 Guardar usuarios"):
            new_users = [u.strip() for u in users_txt.splitlines() if u.strip()]
            if new_users:
                st.session_state.usuarios = new_users
                st.success(f"✅ {len(new_users)} usuario(s) guardado(s)")
            else:
                st.warning("La lista no puede estar vacía")

    with col2:
        st.markdown("**📍 Posición del sello en el PDF**")
        st.caption("El PDF de facturas es horizontal (792 × 612 pts). "
                   "El sello tiene fondo blanco y borde rojo.")

        c_x, c_y = st.columns(2)
        with c_x:
            new_sx = st.number_input("X — desde izquierda", 0, 780,
                                     st.session_state.stamp_x, step=5)
        with c_y:
            new_sy = st.number_input("Y — tope del sello (desde abajo)", 0, 610,
                                     st.session_state.stamp_y_top, step=5)

        c_w, c_h = st.columns(2)
        with c_w:
            new_sw = st.number_input("Ancho", 80, 400, st.session_state.stamp_w, step=5)
        with c_h:
            new_sh = st.number_input("Alto", 40, 200, st.session_state.stamp_h, step=5)

        if st.button("💾 Guardar posición"):
            st.session_state.stamp_x     = new_sx
            st.session_state.stamp_y_top = new_sy
            st.session_state.stamp_w     = new_sw
            st.session_state.stamp_h     = new_sh
            st.success("✅ Posición actualizada")

        st.info(
            "💡 **Valores por defecto (centrado arriba):** X=281, Y=594, Ancho=230, Alto=82  \n"
            "Ajusta si el sello no cae en el espacio correcto de la factura."
        )

    st.divider()

    with st.expander("ℹ️ Información del sistema"):
        st.markdown(f"""
        **Codificador de Facturas v1.0**

        **Lógica de codificación:**
        - Los primeros **2 caracteres** del N° de Orden del cliente determinan el prefijo CC
        - El prefijo se consulta en la **tabla de Proveedores** → Vendor + Centro de Costo
        - El **Código de Producto** se busca en la factura y se consulta en la **tabla GL** → Cuenta GL
        - **Excepción facturas tipo 6:** si el N° de factura comienza con `6` →
          Vendor = `{VENDOR_EXCEPCION}` (fijo), CC = selección manual del usuario

        **Formatos de sello:**
        ```
        POSTED BY: [usuario]
        VENDOR: [número de vendor]
        CC: [centro de costo]  |  GL: [cuenta GL]
        DATE: [DD/MM/YYYY]
        ```

        **Proveedores en base de datos:** {len(st.session_state.proveedores)}
        **Códigos GL en base de datos:** {len(st.session_state.gl_codes)}
        **Usuarios:** {', '.join(st.session_state.usuarios)}
        """)
