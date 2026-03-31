"""
outputs/excel_builder.py
========================
Builds the BMS Attribution Model Excel output workbook.

Sheet structure:
    Model_Index         — build summary and version info
    Cost_Code_Map       — master cost code classification table
    Cluster_Rule        — formal ISELL/ICOST rule documentation
    Attribution_Log     — full audit trail, one row per BMS cluster
    Labour_Summary      — hours and P&L by file and combined

All SUBTOTAL and COMBINED P&L rows use Excel formulas so that
manual adjustments to data rows flow through automatically.
"""

import os
from datetime import date
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from core.rates import format_rate_legend

# ── Style constants ───────────────────────────────────────────────────────────
C_NAV  = "1F3864"; C_MID  = "2E75B6"; C_GRN  = "C6EFCE"; C_LGRN = "E2EFDA"
C_SALM = "FCE4D6"; C_YEL  = "FFF2CC"; C_BLU  = "DEEAF1"; C_GRY  = "F5F5F5"
C_AMB  = "FFEB9C"; C_WHT  = "FFFFFF"

TODAY  = date.today().strftime("%d %B %Y")
VER    = "3.1"


def _fill(h):   return PatternFill("solid", fgColor=h)
def _tb(c="BDD7EE"):
    s = Side(style="thin", color=c)
    return Border(left=s, right=s, top=s, bottom=s)
def _f(sz=9, bold=False, color="000000"):
    return Font(name="Arial", size=sz, bold=bold, color=color)
def _cen():  return Alignment(horizontal="center", vertical="center", wrap_text=True)
def _lft(w=True):
    return Alignment(horizontal="left",   vertical="center", wrap_text=w)
def _rgt():  return Alignment(horizontal="right",  vertical="center")


def _title_bar(ws, row, text, bg, sz=12, cols=16, height=22):
    ws.merge_cells(f"A{row}:{get_column_letter(cols)}{row}")
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(name="Arial", size=sz, bold=True, color=C_WHT)
    c.fill = _fill(bg); c.alignment = _cen(); c.border = _tb()
    ws.row_dimensions[row].height = height


def _hdr_row(ws, row, vals, bg=C_MID, sz=9, height=32):
    for col, v in enumerate(vals, 1):
        c = ws.cell(row=row, column=col, value=v)
        c.font = Font(name="Arial", size=sz, bold=True, color=C_WHT)
        c.fill = _fill(bg); c.alignment = _cen(); c.border = _tb()
    ws.row_dimensions[row].height = height


def _data_row(ws, row, vals, aligns, fmts, bg, bold=False, height=22):
    for col, (val, align, fmt) in enumerate(zip(vals, aligns, fmts), 1):
        c = ws.cell(row=row, column=col, value=val)
        c.font = _f(sz=9, bold=bold,
                    color=C_WHT if col == 1 else "000000")
        c.fill = _fill(C_NAV if col == 1 else bg)
        c.border = _tb()
        c.alignment = (_cen() if align == "c" else
                       _rgt() if align == "r" else _lft())
        if fmt:
            c.number_format = fmt
    ws.row_dimensions[row].height = height


# ── Attribution Log ────────────────────────────────────────────────────────────
def build_attribution_log(wb: Workbook, project_result: dict) -> None:
    ws = wb.create_sheet("Attribution_Log")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "ED7D31"

    pname = project_result["project_name"]
    _title_bar(ws, 1,
        f"PANTA CONTRACTING LIMITED — BMS Attribution Log   |   {pname}   |   v{VER}",
        C_NAV, cols=16)
    _title_bar(ws, 2,
        f"ISELL/ICOST Cluster Scan   |   {TODAY}   |   "
        "Includes Labour Hours & P&L   |   All hours are BMS team hours",
        C_MID, sz=9, cols=16)

    ws.merge_cells("A3:P3")
    n3 = ws["A3"]
    n3.value = (
        "Colour:  Green = BMS cluster rule (512/617)   |   "
        "Amber = Manual flag   |   Yellow = Maintenance (outside revenue scope)   |   "
        "mu_A/B/C = pricing rate bands — all hours are BMS team hours"
    )
    n3.font = _f(sz=9, bold=True, color="7F3F00")
    n3.fill = _fill(C_AMB); n3.alignment = _lft(); n3.border = _tb()
    ws.row_dimensions[3].height = 13

    hdrs = [
        "Log_ID", "Source_File", "Excel_Row", "Item_Number", "Item_Name",
        "BMS_Trigger", "Sell_Total\n(€)", "Mat_Cost\n(€)", "Total\nLab Hrs",
        "mu_A Hrs", "mu_B Hrs", "mu_C Hrs",
        "Lab_Cost\n(€)", "Gross_Profit\n(€)", "Margin\n%", "Hour\nCheck"
    ]
    _hdr_row(ws, 4, hdrs, height=40)

    aligns = ["c","l","c","c","l","c","r","r","r","r","r","r","r","r","c","c"]
    fmts   = [None,None,None,None,None,None,
              '#,##0.00','#,##0.00','#,##0.0',
              '#,##0.0','#,##0.0','#,##0.0',
              '#,##0.00','#,##0.00','0.0%',None]

    DATA_FIRST = 5
    log_row = DATA_FIRST

    for log_id, r in enumerate(project_result["all_clusters"], 1):
        is_flag  = r.get("trigger") == "MANUAL FLAG"
        is_maint = r.get("is_maintenance", False)
        hrs_ok   = r.get("hrs_check", "") == "✓"

        bg = (C_AMB  if is_flag  else
              C_YEL  if is_maint else
              C_LGRN if hrs_ok   else C_SALM)

        margin = (r["gross_profit"] / r["sell_total"] * 100
                  if r["sell_total"] else 0)

        vals = [
            log_id, r["display_name"], r["excel_row"],
            r["item_no"], r["item_name"], r["trigger"],
            r["sell_total"], r["mat_cost"], r["lab_hrs_total"],
            r["mu_A_hrs"], r["mu_B_hrs"], r["mu_C_hrs"],
            r["lab_cost"], r["gross_profit"], margin / 100, r["hrs_check"]
        ]
        _data_row(ws, log_row, vals, aligns, fmts, bg)
        log_row += 1

    DATA_LAST = log_row - 1
    TOT_ROW   = log_row

    # Totals row — SUM formulas
    ws.merge_cells(f"A{TOT_ROW}:F{TOT_ROW}")
    tc = ws.cell(row=TOT_ROW, column=1, value="TOTALS — All BMS Clusters")
    tc.font = _f(sz=10, bold=True, color=C_WHT)
    tc.fill = _fill(C_NAV); tc.alignment = _lft(); tc.border = _tb()

    for col, fmt in [(7,'#,##0.00'),(8,'#,##0.00'),(9,'#,##0.0'),
                     (10,'#,##0.0'),(11,'#,##0.0'),(12,'#,##0.0'),
                     (13,'#,##0.00'),(14,'#,##0.00')]:
        cl = get_column_letter(col)
        c2 = ws.cell(row=TOT_ROW, column=col,
                     value=f"=SUM({cl}{DATA_FIRST}:{cl}{DATA_LAST})")
        c2.font = _f(sz=10, bold=True); c2.fill = _fill(C_GRN)
        c2.alignment = _rgt(); c2.border = _tb(); c2.number_format = fmt

    mg = ws.cell(row=TOT_ROW, column=15,
                 value=f"=IF(G{TOT_ROW}>0,N{TOT_ROW}/G{TOT_ROW},0)")
    mg.font = _f(sz=10, bold=True); mg.fill = _fill(C_GRN)
    mg.alignment = _cen(); mg.border = _tb(); mg.number_format = "0.0%"
    ws.row_dimensions[TOT_ROW].height = 18

    # Column widths
    for i, w in enumerate([8,22,9,14,44,14,13,13,11,10,10,10,13,14,9,12], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A5"


# ── Labour Summary ─────────────────────────────────────────────────────────────
def build_labour_summary(wb: Workbook, project_result: dict) -> None:
    ws = wb.create_sheet("Labour_Summary")
    ws.sheet_view.showGridLines = False
    ws.sheet_properties.tabColor = "70AD47"

    pname = project_result["project_name"]
    _title_bar(ws, 1,
        f"PANTA CONTRACTING LIMITED — BMS Labour Hours & P&L Summary   |   {pname}",
        C_NAV, cols=10)
    _title_bar(ws, 2,
        f"All hours are BMS team hours — mu_A/B/C are pricing rate bands only   |   {TODAY}",
        C_MID, sz=9, cols=10)

    ws.merge_cells("A3:J3")
    n3 = ws["A3"]
    n3.value = (
        "IMPORTANT: The BMS department is entitled to the full revenue, hours and profit "
        "for all clusters shown. mu_A/B/C indicate which labour rate was used for pricing — "
        "they do NOT indicate who performed the work. All work within a BMS cluster is BMS scope."
    )
    n3.font = _f(sz=9, bold=True, color="1F3864")
    n3.fill = _fill(C_BLU); n3.alignment = _lft(); n3.border = _tb()
    ws.row_dimensions[3].height = 24

    DATA_COLS = {
        "C": '#,##0.00', "D": '#,##0.00', "E": '#,##0.0',
        "F": '#,##0.0',  "G": '#,##0.0',  "H": '#,##0.0',
        "I": '#,##0.00', "J": '#,##0.00'
    }

    current_row = 4
    subtotal_rows = []   # track (subtot_row, proj_revenue) for combined block

    for file_res in project_result["files"]:
        grades     = file_res["grades"]
        in_scope   = file_res["in_scope"]
        maint      = file_res["maintenance"]
        all_clust  = in_scope + maint   # maintenance rows go at bottom

        # Section header
        ws.merge_cells(f"A{current_row}:J{current_row}")
        sh = ws.cell(row=current_row, column=1,
                     value=f"  {file_res['display_name']} — {file_res['file_type']}")
        sh.font = _f(sz=10, bold=True, color=C_WHT)
        sh.fill = _fill(C_MID); sh.alignment = _lft(False); sh.border = _tb()
        ws.row_dimensions[current_row].height = 16
        current_row += 1

        # Rate legend
        ws.merge_cells(f"A{current_row}:J{current_row}")
        rl = ws.cell(row=current_row, column=1,
                     value=format_rate_legend(grades))
        rl.font = _f(sz=8, color="595959")
        rl.fill = _fill(C_GRY); rl.alignment = _lft(); rl.border = _tb()
        ws.row_dimensions[current_row].height = 13
        current_row += 1

        # Column headers
        dA = grades.get("mu_A", {}).get("desc", "mu_A")[:10]
        dB = grades.get("mu_B", {}).get("desc", "mu_B")[:10]
        dC = grades.get("mu_C", {}).get("desc", "mu_C")[:10]
        _hdr_row(ws, current_row,
            ["Item", "Item Name", "Sell Total\n(€)", "Mat Cost\n(€)",
             "Lab Hrs\nTotal", f"mu_A\n{dA}", f"mu_B\n{dB}", f"mu_C\n{dC}",
             "Lab Cost\n(€)", "Gross Profit\n(€)"], height=36)
        current_row += 1

        DATA_FIRST_ROW = current_row

        # In-scope rows first
        for r in in_scope:
            is_flag = r.get("is_manual_flag", False)
            bg = C_AMB if is_flag else (C_LGRN if r["hrs_check"] == "✓" else C_SALM)
            name = r["item_name"][:42] + (" [FLAG]" if is_flag else "")
            vals = [r["item_no"], name,
                    r["sell_total"], r["mat_cost"], r["lab_hrs_total"],
                    r["mu_A_hrs"], r["mu_B_hrs"], r["mu_C_hrs"],
                    r["lab_cost"], r["gross_profit"]]
            aligns = ["c","l","r","r","r","r","r","r","r","r"]
            fmts   = [None,None,'#,##0.00','#,##0.00','#,##0.0',
                      '#,##0.0','#,##0.0','#,##0.0','#,##0.00','#,##0.00']
            _data_row(ws, current_row, vals, aligns, fmts, bg, height=14)
            current_row += 1

        DATA_LAST_SCOPE = current_row - 1  # last in-scope row — SUBTOTAL sums up to here

        # SUBTOTAL row — SUM formulas over in-scope rows only
        SUBTOT_ROW = current_row
        ws.cell(row=SUBTOT_ROW, column=1,
                value="SUBTOTAL (in-scope)").font = _f(sz=10, bold=True)
        ws["A" + str(SUBTOT_ROW)].fill = _fill(C_GRN)
        ws["A" + str(SUBTOT_ROW)].alignment = _lft()
        ws["A" + str(SUBTOT_ROW)].border = _tb()

        # Live margin formula in col B
        mb = ws.cell(row=SUBTOT_ROW, column=2,
                     value=(f'=IF(C{SUBTOT_ROW}>0,'
                            f'"Margin "&TEXT(J{SUBTOT_ROW}/C{SUBTOT_ROW},"0.0%"),'
                            f'"Margin —")'))
        mb.font = _f(sz=10, bold=True)
        mb.fill = _fill(C_GRN); mb.alignment = _lft(); mb.border = _tb()

        for col_letter, fmt in DATA_COLS.items():
            c = ws.cell(row=SUBTOT_ROW, column=ord(col_letter) - 64,
                        value=(f"=SUM({col_letter}{DATA_FIRST_ROW}:"
                               f"{col_letter}{DATA_LAST_SCOPE})"))
            c.font = _f(sz=10, bold=True)
            c.fill = _fill(C_GRN); c.alignment = _rgt()
            c.border = _tb(); c.number_format = fmt
        ws.row_dimensions[SUBTOT_ROW].height = 16
        current_row += 1

        # Maintenance rows at bottom (outside subtotal range)
        for r in maint:
            name = r["item_name"][:42] + " [MAINT — excl. from subtotal]"
            vals = [r["item_no"], name,
                    r["sell_total"], r["mat_cost"], r["lab_hrs_total"],
                    r["mu_A_hrs"], r["mu_B_hrs"], r["mu_C_hrs"],
                    r["lab_cost"], r["gross_profit"]]
            aligns = ["c","l","r","r","r","r","r","r","r","r"]
            fmts   = [None,None,'#,##0.00','#,##0.00','#,##0.0',
                      '#,##0.0','#,##0.0','#,##0.0','#,##0.00','#,##0.00']
            _data_row(ws, current_row, vals, aligns, fmts, C_YEL, height=14)
            current_row += 1

        subtotal_rows.append({
            "subtot_row":   SUBTOT_ROW,
            "proj_revenue": file_res["proj_revenue"],
            "display_name": file_res["display_name"],
        })
        current_row += 1   # blank row between sections

    # ── Combined P&L block ────────────────────────────────────────────────────
    # Header
    ws.merge_cells(f"A{current_row}:J{current_row}")
    cb = ws.cell(row=current_row, column=1,
                 value=f"  COMBINED BMS P&L — {pname} (excl. Maintenance)")
    cb.font = Font(name="Arial", size=11, bold=True, color=C_WHT)
    cb.fill = _fill(C_NAV); cb.alignment = _lft(False); cb.border = _tb()
    ws.row_dimensions[current_row].height = 20
    current_row += 1

    # Total project revenue (sum of all file revenues)
    total_proj_rev = sum(
        s["proj_revenue"] for s in subtotal_rows
        if s["proj_revenue"] is not None
    )

    # Build formula references to subtotal rows
    # Each subtotal row has values in cols C–J
    def combined_formula(col_letter):
        parts = [f"{col_letter}{s['subtot_row']}" for s in subtotal_rows]
        return "=" + "+".join(parts)

    START = current_row
    combined_items = [
        # (row_offset, label, formula, fmt, bg, bold)
        (0,  "Combined Project Revenue",
             f"={total_proj_rev}",
             '#,##0.00', C_GRY, False),
        (1,  "BMS Revenue (in-scope)",
             combined_formula("C"),
             '#,##0.00', C_LGRN, True),
        (2,  "BMS Attribution %",
             f"=IF(C{START}>0,C{START+1}/C{START},0)",
             '0.0%', C_LGRN, True),
        (3,  "BMS Material Cost",
             combined_formula("D"),
             '#,##0.00', C_BLU, False),
        (4,  "BMS Labour Hours — Total",
             combined_formula("E"),
             '#,##0.0', C_BLU, False),
        (5,  f"  {project_result['files'][0]['grades'].get('mu_A',{}).get('desc','Electrical')} (mu_A)",
             combined_formula("F"),
             '#,##0.0', C_GRY, False),
        (6,  f"  {project_result['files'][0]['grades'].get('mu_B',{}).get('desc','ECA Tech')} (mu_B)",
             combined_formula("G"),
             '#,##0.0', C_GRY, False),
        (7,  f"  {project_result['files'][0]['grades'].get('mu_C',{}).get('desc','ECA Eng')} (mu_C)",
             combined_formula("H"),
             '#,##0.0', C_GRY, False),
        (8,  "BMS Labour Cost",
             combined_formula("I"),
             '#,##0.00', C_BLU, False),
        (9,  "BMS Gross Profit",
             f"=C{START+1}-C{START+3}-C{START+8}",
             '#,##0.00', C_GRN, True),
        (10, "BMS Gross Margin",
             f"=IF(C{START+1}>0,C{START+9}/C{START+1},0)",
             '0.0%', C_GRN, True),
    ]

    for offset, label, formula, fmt, bg, bold in combined_items:
        row = START + offset
        ws.merge_cells(f"A{row}:B{row}")
        lbl = ws.cell(row=row, column=1, value=label)
        lbl.font = _f(sz=10, bold=bold)
        lbl.fill = _fill(bg); lbl.alignment = _lft(); lbl.border = _tb()

        val = ws.cell(row=row, column=3, value=formula)
        val.font = _f(sz=10, bold=bold)
        val.fill = _fill(bg); val.alignment = _rgt()
        val.border = _tb(); val.number_format = fmt

        for col in range(4, 11):
            c = ws.cell(row=row, column=col)
            c.fill = _fill(bg); c.border = _tb()

        ws.row_dimensions[row].height = 16

    # Column widths
    for i, w in enumerate([10, 42, 13, 13, 11, 10, 10, 10, 13, 13], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"


# ── Main build function ────────────────────────────────────────────────────────
def build_workbook(project_result: dict, output_path: str) -> None:
    """Build the complete output workbook for a project."""
    wb = Workbook()
    wb.remove(wb.active)   # remove default blank sheet

    build_attribution_log(wb, project_result)
    build_labour_summary(wb, project_result)

    wb.save(output_path)
    print(f"\n  Output saved: {output_path}")
