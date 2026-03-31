"""
core/scanner.py
===============
ISELL/ICOST cluster scanner.

Attribution rule:
    - Scan every row of the Elec_Est sheet
    - ISELL (col G / index 6) marks the start of a priced block
    - ICOST (col G / index 6) marks the end of that block
    - If ANY line item within the block has C_CODE 512 or 617 → BMS cluster
    - Entire block sell total attributed to BMS
    - Labour hours extracted from ISELL row col 39 (index)
    - Grade split from d> SUBI rows: col 10=grade, col 19=hours (SUM all d> per grade)
    - Item Number: col D (index 3) on ISELL row
    - Item Name:   col D (index 3) on row immediately below ISELL

Column index reference (0-based):
    3  = col D  (Item Number / Item Name)
    4  = col E  (Sell Rate)
    5  = col F  (Sell Total)
    6  = col G  (ISELL / ICOST marker)
    7  = col H  (C_CODE)
    10 = col K  (Grade label on d> rows)
    17 = col R  (Material cost on d> rows)
    19 = col T  (Labour hours on d> and ISELL rows)
    38 = col AM (Total Material Cost on ISELL row)
    39 = col AN (Total Labour Hours on ISELL row)
"""

import pandas as pd
from config.settings import BMS_ANCHOR_CODES, MAINTENANCE_KEYWORDS, MAINTENANCE_MIN_VALUE
from core.rates import get_rates


def scan_file(filepath: str, display_name: str, sheet: str = "Elec_Est") -> tuple:
    """
    Scan an estimate file for BMS clusters.

    Returns:
        results (list of dicts) — one entry per BMS cluster found
        grades  (dict)          — rate information for this file
    """
    grades = get_rates(filepath)

    try:
        df = pd.read_excel(filepath, sheet_name=sheet, header=None)
    except Exception as e:
        print(f"  [scanner] ERROR reading {filepath} sheet '{sheet}': {e}")
        return [], grades

    # Find all ISELL and ICOST row indices (data rows only, skip header rows 0-12)
    isell_rows = [i for i in range(13, len(df))
                  if str(df.iloc[i, 6]).strip() == "ISELL"]
    icost_rows = [i for i in range(13, len(df))
                  if str(df.iloc[i, 6]).strip() == "ICOST"]

    results = []

    for r in isell_rows:
        # Find the paired ICOST (first ICOST row after this ISELL)
        nxt = next((ic for ic in icost_rows if ic > r), None)
        if nxt is None:
            continue

        # ── Extract item reference ────────────────────────────────────────────
        item_no   = str(df.iloc[r, 3]).strip()
        item_name = str(df.iloc[r + 1, 3]).strip()
        if item_name in ("nan", ""):
            item_name = "—"

        # ── Extract sell values ───────────────────────────────────────────────
        sell_t = df.iloc[r, 5]
        if pd.isna(sell_t):
            sell_t = df.iloc[r, 4]   # qty=1 items: sell_total = sell_rate
        try:
            sell_t = float(sell_t)
        except (TypeError, ValueError):
            sell_t = 0.0

        # ── Extract cost data from ISELL row ──────────────────────────────────
        try:
            mat_cost = float(df.iloc[r, 38])
        except (TypeError, ValueError):
            mat_cost = 0.0

        try:
            lab_hrs = float(df.iloc[r, 39])
        except (TypeError, ValueError):
            lab_hrs = 0.0

        # ── Scan cost codes within cluster ────────────────────────────────────
        codes = []
        for br in range(r + 1, nxt):
            cd = str(df.iloc[br, 7]).strip()
            if cd not in ("nan", "", "C_CODE"):
                codes.append(cd)

        triggers = sorted(set(c for c in codes if c in BMS_ANCHOR_CODES))
        if not triggers:
            continue   # not a BMS cluster

        # ── Extract grade hours — SUM all d> SUBI rows per grade ─────────────
        grade_hrs = {"mu_A": 0.0, "mu_B": 0.0, "mu_C": 0.0}
        for br in range(r + 1, nxt + 1):
            if str(df.iloc[br, 0]).strip() == "d>":
                grade = str(df.iloc[br, 10]).strip()
                if grade in grade_hrs:
                    try:
                        grade_hrs[grade] += float(df.iloc[br, 19])
                    except (TypeError, ValueError):
                        pass

        # ── Labour cost calculation ───────────────────────────────────────────
        lab_cost = sum(
            grade_hrs[g] * grades[g]["cost_rate"]
            for g in ["mu_A", "mu_B", "mu_C"]
            if g in grades
        )
        gross_profit = sell_t - mat_cost - lab_cost

        # ── Hour check ────────────────────────────────────────────────────────
        grade_sum = sum(grade_hrs.values())
        if abs(grade_sum - lab_hrs) < 0.11:
            hrs_check = "✓"
        else:
            hrs_check = f"⚠ sum={grade_sum:.1f} ≠ total={lab_hrs:.1f}"

        # ── Flag maintenance items ────────────────────────────────────────────
        name_lower = item_name.lower()
        is_maintenance = (
            any(kw in name_lower for kw in MAINTENANCE_KEYWORDS)
            and sell_t > MAINTENANCE_MIN_VALUE
        )

        results.append({
            "display_name":  display_name,
            "filepath":      filepath,
            "sheet":         sheet,
            "excel_row":     r + 1,      # 1-based Excel row number
            "item_no":       item_no,
            "item_name":     item_name[:80],
            "trigger":       ",".join(triggers),
            "sell_total":    sell_t,
            "mat_cost":      mat_cost,
            "lab_hrs_total": lab_hrs,
            "mu_A_hrs":      grade_hrs["mu_A"],
            "mu_B_hrs":      grade_hrs["mu_B"],
            "mu_C_hrs":      grade_hrs["mu_C"],
            "lab_cost":      lab_cost,
            "gross_profit":  gross_profit,
            "grades":        grades,
            "hrs_check":     hrs_check,
            "is_maintenance":is_maintenance,
            "is_manual_flag":False,
        })

    return results, grades


def add_manual_flag(flag_def: dict, grades: dict) -> dict:
    """
    Create a manual-flag cluster entry from a settings definition.
    Used for items that are BMS by decision (e.g. MCC-PR) but have no 512/617 code.
    """
    mu_A = flag_def.get("mu_A_hrs", 0.0)
    mu_B = flag_def.get("mu_B_hrs", 0.0)
    mu_C = flag_def.get("mu_C_hrs", 0.0)

    lab_cost = (
        mu_A * grades.get("mu_A", {}).get("cost_rate", 27.46) +
        mu_B * grades.get("mu_B", {}).get("cost_rate", 33.72) +
        mu_C * grades.get("mu_C", {}).get("cost_rate", 30.00)
    )
    sell_t   = flag_def.get("sell_total", 0.0)
    mat_cost = flag_def.get("mat_cost",   0.0)

    return {
        "display_name":  flag_def.get("file", "Manual Flag"),
        "filepath":      flag_def.get("file", ""),
        "sheet":         "Elec_Est",
        "excel_row":     flag_def.get("excel_row", 0),
        "item_no":       flag_def.get("item_no",   "Manual"),
        "item_name":     flag_def.get("item_name", "Manual flag item")[:80],
        "trigger":       "MANUAL FLAG",
        "sell_total":    sell_t,
        "mat_cost":      mat_cost,
        "lab_hrs_total": mu_A + mu_B + mu_C,
        "mu_A_hrs":      mu_A,
        "mu_B_hrs":      mu_B,
        "mu_C_hrs":      mu_C,
        "lab_cost":      lab_cost,
        "gross_profit":  sell_t - mat_cost - lab_cost,
        "grades":        grades,
        "hrs_check":     "⚠ estimated",
        "is_maintenance":False,
        "is_manual_flag":True,
        "flag_reason":   flag_def.get("reason", ""),
    }
