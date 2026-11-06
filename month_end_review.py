"""Month-end review of the 112-000 summary (Atlantic / Mitchel-Lincoln classeur).

Fills the 112-000 summary with formulas, then adds the sheets
'Couru', 'Ajustes proveedor', 'Correcciones factura' and 'Observaciones'.

Used by the "Month-end 112-000" module in app.py; can also run from the command line:
    python month_end_review.py CLASSEUR.xlsx CLASSEUR_revisado.xlsx
"""
import sys, re, collections
from io import BytesIO
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

TAX = 1.14975
GROUPS = [  # statement, system accounts, row of the accrual in 112-000
    ("635108", ["0101000430", "0101002430", "0101002220"], 24),
    ("218005", ["0101000390", "0101000450"], 25),
    ("NT0039", ["0101000440"], 26),
    ("holding", ["0101004440"], 27),
]


def norm(v):
    """Normalise an invoice number: text/number, '409 40914921' and 40940914921 -> '40914921'."""
    if v is None: return None
    s = str(v).strip()
    if not s: return None
    s = s.split()[-1]
    try: s = str(int(float(s)))
    except ValueError: return s
    if len(s) == 11 and s.startswith("409"): s = s[3:]
    return s


def read_lists(src_bytes):
    """Return (adjustments, corrections) already saved in the workbook, as lists of 3-tuples."""
    wb = openpyxl.load_workbook(BytesIO(src_bytes), read_only=True, data_only=True)
    names = {n.strip().lower(): n for n in wb.sheetnames}
    def rows(name, need):
        if name.lower() not in names: return []
        out = []
        for r in wb[names[name.lower()]].iter_rows(min_row=5, values_only=True):
            r = (list(r) + [None] * 3)[:3]
            if all(r[i] not in (None, "") for i in need):
                out.append(tuple("" if v is None else str(v).strip() for v in r))
        return out
    res = rows("Ajustes proveedor", (0, 1)), rows("Correcciones factura", (0, 1))
    wb.close()
    return res


def run_review(src_bytes, adj_rows=(), fix_rows=()):
    """Run the review. adj_rows: [(statement, number, description)]; fix_rows: [(wrong no., right no., description)].
    Returns (xlsx_bytes, summary dict)."""
    vals = openpyxl.load_workbook(BytesIO(src_bytes), data_only=True)
    wb = openpyxl.load_workbook(BytesIO(src_bytes))
    names = {n.strip().lower(): n for n in wb.sheetnames}
    real = lambda n: names.get(str(n).strip().lower())
    sheet = lambda n: vals[real(n)] if real(n) else None
    num = lambda v: isinstance(v, (int, float)) and not isinstance(v, bool)

    def norm(v):
        if v is None: return None
        s = str(v).strip()
        if not s: return None
        s = s.split()[-1]
        try: s = str(int(float(s)))
        except ValueError: return s
        if len(s) == 11 and s.startswith("409"): s = s[3:]
        return s

    adj_list = {(str(g).strip(), norm(n)): str(d or "") for g, n, d in adj_rows if n not in (None, "")}
    fix_list = {norm(w): (norm(r), str(d or "")) for w, r, d in fix_rows if w not in (None, "") and r not in (None, "")}

    # ---------- notas manuales de 112-000 (I/J) ----------
    summ_vals = vals["112-000"]
    manual_notes = {}
    for r in range(5, 33):
        j, i = summ_vals[f"J{r}"].value, summ_vals[f"I{r}"].value
        if isinstance(j, str):
            for m in re.findall(r"\d{7,}", j):
                manual_notes[m] = f"Mencionada en ajuste manual I{r} ({i:,.2f}): {j.strip()}" if num(i) else f"Mencionada en J{r}: {j.strip()}"

    # ---------- facturas registradas ----------
    registered = {}
    for _, accts, _ in GROUPS:
        for a in accts:
            ws = sheet(a)
            if ws is None: continue
            for r in ws.iter_rows(min_row=2, values_only=True):
                if len(r) < 14 or r[4] is None or not num(r[13]): continue
                k = norm(r[4])
                if k in registered: registered[k][1] += r[13]
                else: registered[k] = [a, r[13], r[3], str(r[4]).strip()]

    fixed = {}   # número correcto -> número en sistema
    for wrong, (right, _) in fix_list.items():
        if wrong in registered and right not in registered:
            registered[right] = registered.pop(wrong); fixed[right] = wrong

    def statement(name):
        ws = sheet(name); lines = []; total = None
        if ws is None: return lines, 0
        for r in ws.iter_rows(values_only=True):
            r = list(r) + [None] * 9
            if isinstance(r[1], str) and r[1].strip().lower().startswith("total ow") and num(r[6]):
                total = r[6]
            elif r[0] == "CAD" and num(r[3]) and num(r[6]):
                lines.append(dict(inv=norm(r[3]), date=r[1], ref=(str(r[4]).strip() if r[4] is not None else ""),
                                  amt=r[6], orig=r[5]))
        if total is None: total = sum(l["amt"] for l in lines)
        return lines, total

    # ---------- pagos en circulación ----------
    ap = sheet("AP INVOICE")
    ap_rows = [list(r) + [None] * 28 for r in ap.iter_rows(min_row=2, values_only=True)] if ap else []
    ap_by_inv = {}
    for r in ap_rows: ap_by_inv.setdefault(norm(r[5]), r)
    pays = collections.defaultdict(list)
    for r in ap_rows:
        if r[22]: pays[str(r[22]).strip()].append(r)
    oc = sheet("O-S CHQS")
    paid, pay_summary = {}, []
    for c in range(1, oc.max_column + 1):
        vend = oc.cell(1, c).value
        if vend is None or not str(vend).strip(): continue
        vend = str(vend).strip()
        for rr in range(2, 7):
            a = oc.cell(rr, c).value
            if not num(a) or abs(a) < 0.005: continue
            found = next((p for p, rows in pays.items()
                          if any(str(x[13]).strip() == vend for x in rows)
                          and any(num(x[25]) and abs(x[25] - a) < 0.01 for x in rows)), None)
            invs = {}
            if found:
                for x in pays[found]: invs[norm(x[5])] = x[20]
                for k, t in invs.items(): paid[k] = dict(pay=found, total=t, vend=vend)
            pay_summary.append(dict(vend=vend, amt=a, pay=found, date=pays[found][0][23] if found else None,
                                    n=len(invs), invsum=sum(v for v in invs.values() if num(v))))

    def ap_note(k):
        r = ap_by_inv.get(k)
        if not r: return ""
        s = f"AP INVOICE: cuenta {r[13]}" + (f", total {r[20]:,.2f}" if num(r[20]) else "")
        if r[22]: s += f", pago {r[22]} ({r[24]})"
        return s

    # ---------- clasificación ----------
    GN = [g for g, _, _ in GROUPS]
    group_of = {a: g for g, accts, _ in GROUPS for a in accts}
    detail, on_stmt = [], set()
    reclass = collections.defaultdict(float); reclass_invs = collections.defaultdict(list)
    candidates = []
    for g, accts, _ in GROUPS:
        lines, _ = statement(g)
        for l in lines:
            k = l["inv"]; on_stmt.add(k)
            row = dict(stmt=g, inv=k, date=l["date"], ref=l["ref"], amt=l["amt"], acct="", sysamt=None, pay="", note="")
            if (g, k) in adj_list:
                row.update(cat="Ajuste del proveedor (no es factura)", note=adj_list[(g, k)])
                refk = norm(l["ref"])
                for sk, (sa, samt, _, sref) in registered.items():
                    if sk not in on_stmt and refk and sk.startswith(refk) and abs(samt - l["amt"]) < 0.01:
                        row.update(acct=sa, sysamt=samt, note=row["note"] + f"; compensa {sref} del sistema")
                        row["pair"] = sk
            elif k in registered:
                a, sa, _, _ = registered[k]
                row.update(acct=a, sysamt=sa)
                if a in accts:
                    row["cat"] = "Registrada"
                    if k in fixed:
                        row["note"] = f"Registrada en el sistema con número erróneo {fixed[k]} ({fix_list[fixed[k]][1]})"
                else:
                    row["cat"] = "Registrada en otra cuenta"
                    row["note"] = f"Registrada en {a}; reclasificar a {accts[0]}"
                    reclass[(a, accts[0])] += sa; reclass_invs[(a, accts[0])].append(k)
                if abs(sa - l["amt"]) > 0.01: row["note"] = (row["note"] + "; " if row["note"] else "") + "Monto distinto entre sistema y estado"
            elif k in paid:
                p = paid[k]
                row.update(cat="Pagada - cheque en circulación", acct=p["vend"], sysamt=p["total"], pay=p["pay"])
                if num(p["total"]) and abs(p["total"] - l["amt"]) > 0.01: row["note"] = "Monto pagado distinto al del estado"
            else:
                row["cat"] = "No registrada"
                row["note"] = "; ".join(x for x in (ap_note(k), manual_notes.get(k, "")) if x)
                # ¿parece un ajuste y no una factura?
                why = []
                if "EFT" in l["ref"].upper(): why.append("referencia EFT")
                if num(l["orig"]) and abs(l["orig"] - l["amt"]) > 0.01: why.append("monto original ≠ saldo")
                rk = norm(l["ref"])
                if rk and rk != k and (rk in registered or rk in ap_by_inv or rk in paid): why.append(f"referencia es otra factura ({l['ref']})")
                if why: candidates.append((g, k, l["amt"], ", ".join(why)))
            detail.append(row)
    paired = {r["pair"] for r in detail if r.get("pair")}
    for k, (a, sa, d, ref) in registered.items():
        if k not in on_stmt:
            detail.append(dict(stmt=group_of.get(a, ""), cat="En sistema, no en estado de cuenta", inv=ref, date=d, ref="",
                               amt=None, acct=a, sysamt=sa, pay="",
                               note=("Compensada con ajuste del proveedor; " if k in paired else "") + ap_note(k)))
    for k, p in paid.items():
        if k not in on_stmt and k not in registered:
            detail.append(dict(stmt=group_of.get(p["vend"], ""), cat="Pagada, no en estado de cuenta", inv=k, date=None, ref="",
                               amt=None, acct=p["vend"], sysamt=p["total"], pay=p["pay"], note="En pago en circulación pero no en el estado"))
    ORDER = ["No registrada", "Ajuste del proveedor (no es factura)", "Registrada en otra cuenta",
             "En sistema, no en estado de cuenta", "Pagada, no en estado de cuenta", "Pagada - cheque en circulación", "Registrada"]
    detail.sort(key=lambda r: (ORDER.index(r["cat"]), GN.index(r["stmt"]) if r["stmt"] in GN else 9, str(r["inv"])))
    unreg = [r for r in detail if r["cat"] == "No registrada"]

    # ---------- estilos ----------
    F = "Arial"
    H = Font(name=F, bold=True, color="FFFFFF", size=10); HF = PatternFill("solid", fgColor="1F4E78")
    B = Font(name=F, bold=True, size=10); N = Font(name=F, size=10); INP = Font(name=F, size=10, color="0000FF")
    T = Font(name=F, bold=True, size=14); S = Font(name=F, bold=True, size=11, color="1F4E78")
    SUB = Font(name=F, italic=True, size=9, color="595959")
    WARN = PatternFill("solid", fgColor="FCE4D6"); YEL = PatternFill("solid", fgColor="FFF2CC")
    MONEY = '#,##0.00;(#,##0.00);"-"'
    def header(ws, row, labels):
        for i, h in enumerate(labels, 1):
            c = ws.cell(row, i, h); c.font = H; c.fill = HF; c.alignment = Alignment(wrap_text=True, vertical="center")
        ws.row_dimensions[row].height = 30
    def widths(ws, ws_w):
        for i, w in enumerate(ws_w, 1): ws.column_dimensions[get_column_letter(i)].width = w

    # keep edits already made in a previous Couru sheet (tax factor / Incluir)
    prev_couru = {}
    if sheet("Couru") is not None:
        for r in sheet("Couru").iter_rows(values_only=True):
            r = list(r) + [None] * 9
            if r[0] and r[1] and num(r[5]) and r[7] in ("Sí", "No"):
                prev_couru[(str(r[0]).strip(), norm(r[1]))] = (r[5], r[7])
    for n in ("Couru", "Ajustes proveedor", "Correcciones factura", "Observaciones"):
        if n in wb.sheetnames: del wb[n]

    # ---------- hoja Couru ----------
    cu = wb.create_sheet("Couru")
    cu["A1"] = "Couru – facturas no contabilizadas"; cu["A1"].font = T
    cu["A2"] = ("Facturas de los estados de cuenta que no están registradas ni pagadas en el sistema. "
                "Editar en azul: factor de tax (por defecto 1.14975) e 'Incluir' (Sí/No). Las provisiones D24:E27 del 112-000 salen de aquí.")
    cu["A2"].font = SUB
    header(cu, 4, ["Estado de cuenta", "Cuenta 112-000", "N° facturas incluidas", "Monto con taxes", "Neto (sin taxes)", "Taxes"])
    CF = 9 + len(GROUPS)           # encabezado del detalle
    c1, c2 = CF + 1, CF + max(len(unreg), 1)
    R = lambda col: f"${col}${c1}:${col}${c2}"
    for j, (g, _, prow) in enumerate(GROUPS):
        r = 5 + j
        cu.cell(r, 1, g); cu.cell(r, 2, f"D{prow} / E{prow}")
        cu.cell(r, 3, f'=COUNTIFS({R("A")},A{r},{R("H")},"Sí")')
        cu.cell(r, 4, f'=ROUND(SUMIFS({R("E")},{R("A")},A{r},{R("H")},"Sí"),2)')
        cu.cell(r, 5, f'=ROUND(SUMIFS({R("G")},{R("A")},A{r},{R("H")},"Sí"),2)')
        cu.cell(r, 6, f"=D{r}-E{r}")
    tr = 5 + len(GROUPS)
    cu.cell(tr, 1, "TOTAL").font = B
    for col in "CDEF":
        cu[f"{col}{tr}"] = f"=SUM({col}5:{col}{tr-1})"; cu[f"{col}{tr}"].font = B
    for r in range(5, tr + 1):
        for col in "DEF": cu[f"{col}{r}"].number_format = MONEY
    cu.cell(CF - 1, 1, "Detalle").font = S
    header(cu, CF, ["Estado de cuenta", "Factura", "Fecha factura", "Referencia cliente", "Monto con taxes",
                    "Factor tax", "Neto", "Incluir", "Nota"])
    for j, d in enumerate(unreg):
        r = c1 + j
        cu.cell(r, 1, d["stmt"]); cu.cell(r, 2, d["inv"])
        if d["date"] is not None: cu.cell(r, 3, d["date"]).number_format = "yyyy-mm-dd"
        cu.cell(r, 4, d["ref"]); cu.cell(r, 5, d["amt"]).number_format = MONEY
        prev = prev_couru.get((d["stmt"], d["inv"]))
        f = cu.cell(r, 6, prev[0] if prev else TAX); f.font = INP; f.number_format = "0.00000"
        cu.cell(r, 7, f"=ROUND(E{r}/F{r},2)").number_format = MONEY
        inc = cu.cell(r, 8, prev[1] if prev else ("No" if d["inv"] in manual_notes else "Sí")); inc.font = INP; inc.alignment = Alignment(horizontal="center")
        cu.cell(r, 9, d["note"])
        if d["inv"] in manual_notes or any(c[1] == d["inv"] for c in candidates):
            for col in range(1, 10): cu.cell(r, col).fill = YEL
    dv = DataValidation(type="list", formula1='"Sí,No"', allow_blank=False); cu.add_data_validation(dv); dv.add(f"H{c1}:H{c2}")
    cu.auto_filter.ref = f"A{CF}:I{c2}"
    widths(cu, [16, 14, 13, 26, 17, 11, 16, 9, 70])
    cu.row_dimensions[2].height = 26; cu["A2"].alignment = Alignment(wrap_text=True); cu.merge_cells("A2:I2")
    cu.freeze_panes = f"A{c1}"

    # ---------- hoja Ajustes proveedor ----------
    aj = wb.create_sheet("Ajustes proveedor")
    aj["A1"] = "Ajustes del proveedor (no son facturas)"; aj["A1"].font = T
    aj["A2"] = ("Líneas de los estados de cuenta que son ajustes y no facturas: no entran en el couru. "
                "Agregar aquí nuevas líneas cada mes (estado de cuenta + número).")
    aj["A2"].font = SUB
    header(aj, 4, ["Estado de cuenta", "Número en el estado", "Descripción", "Monto en el estado (mes actual)"])
    amt_by = {(r["stmt"], r["inv"]): r["amt"] for r in detail if r["cat"] == "Ajuste del proveedor (no es factura)"}
    for j, ((g, k), desc) in enumerate(adj_list.items()):
        r = 5 + j
        for col, v in enumerate([g, k, desc], 1):
            c = aj.cell(r, col, v); c.font = INP
        m = amt_by.get((g, k))
        aj.cell(r, 4, m if m is not None else "No aparece este mes").number_format = MONEY
    widths(aj, [16, 20, 70, 22])

    pts = []
    for (a, b), m in reclass.items():
        pts.append(f"Reclasificación: {', '.join(reclass_invs[(a,b)])} registradas en {a} aparecen en el estado {group_of[b]} → "
                   f"{m:,.2f} pasan de {a} a {b} (aplicado en columna E del resumen).")
    for right, wrong in fixed.items():
        pts.append(f"Corrección de número: {wrong} en el sistema = {right} en el estado ({registered[right][1]:,.2f}); tratada como registrada.")
    for wrong in fix_list:
        if wrong not in fixed.values():
            pts.append(f"Corrección {wrong} → {fix_list[wrong][0]} de la hoja 'Correcciones factura' no aplica este mes (ya no está en el sistema o el número correcto ya existe).")
    pts.append(f"{len(unreg)} facturas no contabilizadas → hoja 'Couru' (base de D24:E27).")
    for r in detail:
        if r["cat"] == "Ajuste del proveedor (no es factura)":
            pts.append(f"Ajuste del proveedor {r['inv']} ({r['stmt']}) {r['amt']:,.2f} excluido del couru: {r['note']}.")
    so = [r for r in detail if r["cat"] == "En sistema, no en estado de cuenta" and not r["note"].startswith("Compensada")]
    cancel = set()
    for i, a_ in enumerate(so):
        for b_ in so[i+1:]:
            if a_["acct"] == b_["acct"] and abs(a_["sysamt"] + b_["sysamt"]) < 0.01 and id(a_) not in cancel and id(b_) not in cancel:
                cancel |= {id(a_), id(b_)}
                a_["note"] = f"Se anula con {b_['inv']}; " + a_["note"]; b_["note"] = f"Se anula con {a_['inv']}; " + b_["note"]
                pts.append(f"En sistema y no en el estado, se anulan entre sí (neto 0): {a_['inv']} y {b_['inv']} ({a_['acct']}).")
    so = [r for r in so if id(r) not in cancel]
    for r in so:
        same = [u for u in unreg if u["stmt"] == r["stmt"] and abs(u["amt"] - r["sysamt"]) < 0.01]
        if same:
            pts.append(f"¿Número erróneo? {r['inv']} en el sistema y {', '.join(u['inv'] for u in same)} en el estado tienen el mismo monto "
                       f"({r['sysamt']:,.2f}). Si es la misma factura, agregarla en 'Correcciones factura'.")
    for r in so:
        pts.append(f"En sistema y no en el estado: {r['inv']} ({r['acct']}) {r['sysamt']:,.2f}" + (f" – {r['note']}" if r["note"] else ""))
    for p in pay_summary:
        if not p["pay"]: pts.append(f"Pago de {p['amt']:,.2f} de {p['vend']} en O-S CHQS no encontrado en AP INVOICE.")
        elif abs(p["amt"] - p["invsum"]) > 0.01:
            pts.append(f"Pago {p['pay']}: monto O-S CHQS {p['amt']:,.2f} vs facturas del pago {p['invsum']:,.2f} "
                       f"(diferencia {p['amt']-p['invsum']:,.2f}).")
    for g, k, a, why in candidates:
        pts.append(f"¿Ajuste y no factura? {k} ({g}) {a:,.2f}: {why}. Si lo es, agregarlo en 'Ajustes proveedor'.")
    for k, t in manual_notes.items():
        if any(r["inv"] == k for r in unreg): pts.append(f"Factura {k} marcada 'No' en el Couru porque está en tus ajustes manuales: {t}. Cambiar a 'Sí' si debe provisionarse.")
    amtd = [r for r in detail if "distinto" in r["note"]]
    pts.append(f"{len(amtd)} facturas con monto distinto entre sistema y estado." if amtd else
               "Ninguna factura registrada tiene monto distinto entre sistema y estado de cuenta.")
    # ---------- hoja Correcciones factura ----------
    cf = wb.create_sheet("Correcciones factura")
    cf["A1"] = "Correcciones de número de factura"; cf["A1"].font = T
    cf["A2"] = ("Facturas registradas en el sistema con un número erróneo. Se tratan como registradas con el número correcto "
                "(el que aparece en el estado de cuenta). Agregar nuevas cada mes; las que ya no apliquen se pueden borrar.")
    cf["A2"].font = SUB
    header(cf, 4, ["Número en el sistema", "Número correcto (estado)", "Descripción", "Monto (mes actual)"])
    for j, (wrong, (right, desc)) in enumerate(fix_list.items()):
        r = 5 + j
        for col, v in enumerate([wrong, right, desc], 1):
            c = cf.cell(r, col, v); c.font = INP
        cf.cell(r, 4, registered[right][1] if fixed.get(right) == wrong else "No aplica este mes").number_format = MONEY
    widths(cf, [22, 24, 70, 20])

    # ---------- hoja Observaciones ----------
    ob = wb.create_sheet("Observaciones")
    ob["A1"] = "Observaciones – revisión 112-000"; ob["A1"].font = T
    ob["A2"] = ("Comparación de cada estado de cuenta contra las facturas registradas (hojas 0101…) y los pagos en circulación "
                "(O-S CHQS / AP INVOICE).")
    ob["A2"].font = SUB
    sum_start = 5
    pay_start = sum_start + len(GROUPS) + 4
    chk_start = pay_start + len(pay_summary) + 4
    DT = chk_start + len(pts) + 2
    first, last = DT + 1, DT + len(detail)
    rng = lambda col: f"${col}${first}:${col}${last}"
    ob.cell(sum_start - 1, 1, "1. Cuadre por estado de cuenta").font = S
    cats = ["Registrada", "Registrada en otra cuenta", "Pagada - cheque en circulación", "Ajuste del proveedor (no es factura)", "No registrada"]
    header(ob, sum_start, ["Estado de cuenta", "Cuentas sistema", "Total estado (Total Ow)", "Registradas", "Registradas en otra cuenta",
                           "Pagadas – cheque en circulación", "Ajustes del proveedor", "No registradas (con taxes)", "Control (debe ser 0)"])
    ob.row_dimensions[sum_start].height = 42
    for j, (g, accts, _) in enumerate(GROUPS):
        r = sum_start + 1 + j
        ob.cell(r, 1, g); ob.cell(r, 2, ", ".join(accts))
        sg = real(g)
        if sg and g in ("635108", "218005"):
            ob.cell(r, 3, f"=IFERROR(INDEX('{sg}'!G:G,MATCH(\"Total Ow*\",'{sg}'!B:B,0)),0)")
        else:
            ob.cell(r, 3, f"=SUM('{sg}'!G:G)" if sg else 0)
        for col, cat in zip("DEFGH", cats):
            ob[f"{col}{r}"] = f'=SUMIFS({rng("F")},{rng("A")},$A{r},{rng("B")},"{cat}")'
        ob[f"I{r}"] = f"=ROUND(C{r}-SUM(D{r}:H{r}),2)"
    tr = sum_start + 1 + len(GROUPS)
    ob.cell(tr, 1, "TOTAL").font = B
    for col in "CDEFGHI":
        ob[f"{col}{tr}"] = f"=SUM({col}{sum_start+1}:{col}{tr-1})"; ob[f"{col}{tr}"].font = B
    for r in range(sum_start + 1, tr + 1):
        for col in "CDEFGHI": ob[f"{col}{r}"].number_format = MONEY

    ob.cell(pay_start - 1, 1, "2. Pagos en circulación (O-S CHQS) encontrados en AP INVOICE").font = S
    header(ob, pay_start, ["Cuenta (vendor)", "Monto O-S CHQS", "N° de pago", "Fecha de pago", "N° facturas",
                           "Total facturas del pago", "Diferencia", "Facturas del pago en el estado"])
    for j, p in enumerate(pay_summary):
        r = pay_start + 1 + j
        ob.cell(r, 1, p["vend"]); ob.cell(r, 2, p["amt"]).number_format = MONEY
        ob.cell(r, 3, p["pay"] or "NO ENCONTRADO")
        if p["date"]: ob.cell(r, 4, p["date"]).number_format = "yyyy-mm-dd"
        ob.cell(r, 5, p["n"]); ob.cell(r, 6, p["invsum"]).number_format = MONEY
        ob.cell(r, 7, f"=ROUND(B{r}-F{r},2)").number_format = MONEY
        ob.cell(r, 8, f'=SUMIFS({rng("F")},{rng("K")},C{r})').number_format = MONEY
        if not p["pay"]: ob.cell(r, 3).fill = WARN

    # puntos a revisar
    ob.cell(chk_start - 1, 1, "3. Puntos a revisar").font = S
    for i, t in enumerate(pts):
        c = ob.cell(chk_start + i, 1, "• " + t); c.font = N; c.alignment = Alignment(wrap_text=True, vertical="top")
        ob.merge_cells(start_row=chk_start + i, start_column=1, end_row=chk_start + i, end_column=9)
        ob.row_dimensions[chk_start + i].height = 27

    ob.cell(DT - 1, 1, "4. Detalle factura por factura (usar el filtro de 'Categoría')").font = S
    header(ob, DT, ["Estado de cuenta", "Categoría", "Factura", "Fecha factura", "Referencia cliente", "Monto estado (con taxes)",
                    "—", "Cuenta sistema", "Monto sistema", "Diferencia", "N° de pago", "Nota"])
    ob.cell(DT, 7, "")
    for j, d in enumerate(detail):
        r = first + j
        ob.cell(r, 1, d["stmt"]); ob.cell(r, 2, d["cat"]); ob.cell(r, 3, d["inv"])
        if d["date"] is not None: ob.cell(r, 4, d["date"]).number_format = "yyyy-mm-dd"
        ob.cell(r, 5, d["ref"])
        if d["amt"] is not None: ob.cell(r, 6, d["amt"])
        ob.cell(r, 8, d["acct"])
        if d["sysamt"] is not None: ob.cell(r, 9, d["sysamt"])
        if d["amt"] is not None and d["sysamt"] is not None: ob.cell(r, 10, f"=ROUND(F{r}-I{r},2)")
        ob.cell(r, 11, d["pay"]); ob.cell(r, 12, d["note"])
        for col in (6, 9, 10): ob.cell(r, col).number_format = MONEY
        if d["cat"] not in ("Registrada", "Pagada - cheque en circulación") or "distinto" in d["note"] or "erróneo" in d["note"]:
            for col in range(1, 13): ob.cell(r, col).fill = WARN
    ob.auto_filter.ref = f"A{DT}:L{last}"
    widths(ob, [16, 32, 14, 13, 26, 18, 3, 16, 16, 13, 13, 60])
    ob.sheet_view.showGridLines = False
    for wsx in (cu, aj, cf, ob):
        for row in wsx.iter_rows(min_row=3):
            for c in row:
                if c.value is not None and (c.font is None or c.font.name != F):
                    c.font = Font(name=F, size=10, bold=c.font.bold if c.font else False,
                                  color=c.font.color if c.font and c.font.color else None)

    # ---------- resumen 112-000 ----------
    sm = wb["112-000"]
    adj = collections.defaultdict(float); adj_txt = collections.defaultdict(list)
    for (a, b), m in reclass.items():
        lst = ", ".join(reclass_invs[(a, b)])
        adj[a] -= m; adj[b] += m
        adj_txt[a].append(f"−{m:,.2f} reclasificado a {b} (facturas {lst})")
        adj_txt[b].append(f"+{m:,.2f} reclasificado desde {a} (facturas {lst})")
    for r in range(6, 13):
        acct = str(sm[f"D{r}"].value or "").strip(); s = real(acct)
        base = f"SUM('{s}'!N:N)" if s else "0"
        a = round(adj.get(acct, 0), 2)
        sm[f"E{r}"] = f"={base}{a:+.2f}" if a else f"={base}"
        sm[f"E{r}"].comment = None
        if a: sm[f"E{r}"].comment = Comment("Reclasificación automática: " + "; ".join(adj_txt[acct]) + ". Ver Observaciones.", "Revisión")
        elif not s: sm[f"E{r}"].comment = Comment(f"No existe hoja {acct}: saldo 0.", "Revisión")
    s635, s218, snt, shol = real("635108"), real("218005"), real("NT0039"), real("holding")
    sm["I6"] = f"=IFERROR(INDEX('{s635}'!G:G,MATCH(\"Total Ow*\",'{s635}'!B:B,0)),0)" if s635 else 0
    sm["I7"] = f"=IFERROR(INDEX('{s218}'!G:G,MATCH(\"Total Ow*\",'{s218}'!B:B,0)),0)" if s218 else 0
    sm["I8"] = f"=SUM('{snt}'!G:G)" if snt else 0
    sm["I9"] = f"=SUM('{shol}'!G:G)" if shol else 0
    for a in ("I7", "I8", "I9"): sm[a].number_format = sm["I6"].number_format
    for r in range(15, 22):
        sm[f"E{r}"] = f"=IFERROR(INDEX('O-S CHQS'!$7:$7,MATCH(TRIM(D{r}),'O-S CHQS'!$1:$1,0)),0)"
    for j, (g, _, prow) in enumerate(GROUPS):
        cr = 5 + j
        sm[f"D{prow}"] = f"=Couru!E{cr}"
        sm[f"E{prow}"] = f"=Couru!D{cr}"
        sm[f"D{prow}"].comment = Comment(f"Neto de las facturas no contabilizadas del estado {g} marcadas 'Sí' en la hoja Couru.", "Revisión")
        sm[f"E{prow}"].comment = None
    wb.move_sheet("Couru", offset=0)

    wb.calculation.fullCalcOnLoad = True
    out = BytesIO(); wb.save(out)
    cuadre = []
    for g, accts, _ in GROUPS:
        tot = statement(g)[1]
        by = collections.defaultdict(float)
        for r in detail:
            if r["stmt"] == g and r["amt"] is not None: by[r["cat"]] += r["amt"]
        cuadre.append(dict(statement=g, total=round(tot, 2), **{k: round(v, 2) for k, v in by.items()}))
    summary = dict(
        counts=dict(collections.Counter(r["cat"] for r in detail)),
        cuadre=cuadre,
        couru=[dict(statement=g, invoices=sum(1 for u in unreg if u["stmt"] == g),
                    gross=round(sum(u["amt"] for u in unreg if u["stmt"] == g), 2)) for g, _, _ in GROUPS],
        payments=pay_summary,
        points=pts,
    )
    return out.getvalue(), summary


if __name__ == "__main__":
    data, info = run_review(open(sys.argv[1], "rb").read(), *read_lists(open(sys.argv[1], "rb").read()))
    open(sys.argv[2], "wb").write(data)
    for p in info["points"]: print("-", p)
