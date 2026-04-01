"""
core/attribution.py
===================
Orchestrates the full attribution process for a single project.

Takes a project definition from config/settings.py, runs the scanner
on each file, applies manual flags, and returns a structured result
ready for the Excel builder.
"""

import os
from core.scanner import scan_file, add_manual_flag
from core.rates import get_project_revenue
from config.settings import PROJECTS_DIR


def run_project(project: dict) -> dict:
    """
    Run the full attribution scan for one project.

    Returns a dict containing:
        project_name    : str
        job_number      : str
        files           : list of per-file result dicts
        all_clusters    : flat list of all BMS clusters (all files combined)
        project_revenues: dict of filename -> project total revenue
        notes           : str
    """
    folder      = os.path.join(PROJECTS_DIR, project["folder"])
    revenues    = project.get("project_revenue", {})
    manual_flags= project.get("manual_flags", {})
    notes       = project.get("notes", "")

    file_results = []
    all_clusters = []

    for file_def in project["files"]:
        filename     = file_def["path"]
        filepath     = os.path.join(folder, filename)
        display_name = file_def["label"]
        file_type    = file_def["type"]   # "EL" or "ELM"
        proj_revenue = revenues.get(filename, None)

        # Auto-extract from ML_Sum if not hardcoded in settings
        if proj_revenue is None:
            proj_revenue = get_project_revenue(filepath)
            if proj_revenue:
                print(f"    [revenue] Auto-extracted from ML_Sum: €{proj_revenue:,.2f}")

        print(f"\n  Scanning: {filename} ({file_type})")

        if not os.path.exists(filepath):
            print(f"    [WARNING] File not found: {filepath}")
            continue

        results, grades = scan_file(filepath, display_name)

        # Inject manual flags for this file
        for excel_row, flag_def in manual_flags.items():
            if flag_def.get("file") == filename:
                flag_entry = add_manual_flag(
                    {**flag_def, "excel_row": excel_row},
                    grades
                )
                flag_entry["display_name"] = display_name
                results.append(flag_entry)
                print(f"    [FLAG] Added manual flag: {flag_def.get('item_no')} "
                      f"— {flag_def.get('item_name', '')[:50]}")

        # Separate in-scope and maintenance clusters
        in_scope    = [r for r in results if not r["is_maintenance"]]
        maintenance = [r for r in results if r["is_maintenance"]]

        # Calculate file-level totals
        def totals(cluster_list):
            return {
                "sell":  sum(r["sell_total"]    for r in cluster_list),
                "mat":   sum(r["mat_cost"]      for r in cluster_list),
                "hrs":   sum(r["lab_hrs_total"] for r in cluster_list),
                "mu_A":  sum(r["mu_A_hrs"]      for r in cluster_list),
                "mu_B":  sum(r["mu_B_hrs"]      for r in cluster_list),
                "mu_C":  sum(r["mu_C_hrs"]      for r in cluster_list),
                "lc":    sum(r["lab_cost"]       for r in cluster_list),
                "gp":    sum(r["gross_profit"]   for r in cluster_list),
            }

        scope_tots = totals(in_scope)
        bms_pct    = (scope_tots["sell"] / proj_revenue * 100) if proj_revenue else None
        margin     = (scope_tots["gp"] / scope_tots["sell"] * 100
                      if scope_tots["sell"] else 0)

        # Hour check summary
        bad_checks = [r for r in in_scope if r["hrs_check"] != "OK"]

        print(f"    BMS clusters found:  {len(in_scope)} in-scope, "
              f"{len(maintenance)} maintenance")
        print(f"    BMS revenue:         €{scope_tots['sell']:,.2f}")
        print(f"    BMS attribution:     {bms_pct:.1f}%" if bms_pct else
              f"    BMS attribution:     (no revenue denominator set)")
        print(f"    Gross margin:        {margin:.1f}%")
        print(f"    Labour hours:        {scope_tots['hrs']:,.1f} hrs")
        print(f"    Hour checks:         "
              f"{'ALL OK' if not bad_checks else str(len(bad_checks)) + ' issues'}")
        if bad_checks:
            for r in bad_checks:
                print(f"      [WARNING] {r['item_no']} {r['hrs_check']}")

        file_results.append({
            "filename":      filename,
            "display_name":  display_name,
            "file_type":     file_type,
            "filepath":      filepath,
            "grades":        grades,
            "clusters":      results,          # all clusters (incl maintenance)
            "in_scope":      in_scope,
            "maintenance":   maintenance,
            "scope_totals":  scope_tots,
            "proj_revenue":  proj_revenue,
            "bms_pct":       bms_pct,
            "margin":        margin,
        })

        all_clusters.extend(results)

    return {
        "job_number":       project["job_number"],
        "project_name":     project["project_name"],
        "folder":           folder,
        "files":            file_results,
        "all_clusters":     all_clusters,
        "project_revenues": revenues,
        "notes":            notes,
    }