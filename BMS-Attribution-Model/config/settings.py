"""
BMS Attribution Model — Project Configuration
==============================================
Add new projects to the PROJECTS list.
Each project entry tells the scanner which files to process
and how to label them in the output.

Fields:
    job_number      : Panta job number (e.g. J01577)
    project_name    : Human-readable project name
    folder          : Subfolder name under projects/
    files           : List of estimate files to scan.
                      Each file needs:
                        path    : filename relative to the project folder
                        type    : "EL" or "ELM"
                        label   : short display label for the output
    project_revenue : Total project revenue from SUMM PG (EUR, excl VAT)
                      Used as the denominator for BMS Attribution %.
                      Set to None to use the sum of all ISELL blocks.
    exclude_sheets  : List of sheet names to skip (optional)
    notes           : Any project-specific notes
"""

import os

# ── Base paths ────────────────────────────────────────────────────────────────
# Update BASE_DIR if you move this project to a different machine or the NAS
BASE_DIR     = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PROJECTS_DIR = os.path.join(BASE_DIR, "projects")
OUTPUTS_DIR  = PROJECTS_DIR   # output goes into each project's subfolder

# ── Global settings ───────────────────────────────────────────────────────────
OUTPUT_FILENAME_TEMPLATE = "BMS_Attribution_{job_number}_{project_slug}.xlsx"
CURRENCY = "EUR"

# ── BMS attribution constants ──────────────────────────────────────────────────
BMS_ANCHOR_CODES = {"512", "617"}   # triggers BMS cluster attribution

# Maintenance items — excluded from revenue denominator and subtotals
# Items where item_name contains any of these strings AND sell_total > threshold
MAINTENANCE_KEYWORDS  = ["maintenance", "preventative maint", "prev. maint"]
MAINTENANCE_MIN_VALUE = 5000.0   # only flag as maintenance if value > this

# ── Labour rate fallback defaults ──────────────────────────────────────────────
# Used if rates cannot be extracted from the estimate file
DEFAULT_RATES = {
    "mu_A": {"desc": "Electrical",      "cost_rate": 27.46},
    "mu_B": {"desc": "ECA Technician",  "cost_rate": 33.72},
    "mu_C": {"desc": "ECA Engineer",    "cost_rate": 30.00},
}

# ── Projects ───────────────────────────────────────────────────────────────────
PROJECTS = [
    {
        "job_number":      "J01577",
        "project_name":    "Intercontinental Hotel — Plant Retrofit",
        "folder":          "intercontinental",
        "files": [
            {
                "path":  "E25_0477_IH_Retro_Elec_Est_Rev_A_AW.xlsx",
                "type":  "EL",
                "label": "E25-0477 (EL)",
            },
            {
                "path":  "E25_0465_IH_Retro_Elec4Mech_Est_Rev_A_AW.xlsx",
                "type":  "ELM",
                "label": "E25-0465 (ELM)",
            },
        ],
        # Revenue from SUMM PG sheets (excl VAT, excl discount where applicable)
        "project_revenue": {
            "E25_0477_IH_Retro_Elec_Est_Rev_A_AW.xlsx": 386990.92,
            "E25_0465_IH_Retro_Elec4Mech_Est_Rev_A_AW.xlsx": 164843.55,
        },
        # Manual flag items — BMS attribution by decision, not by 512/617 rule
        # Format: {excel_row: {item_no, item_name, sell_total, mat_cost,
        #                      mu_A_hrs, mu_B_hrs, mu_C_hrs, reason}}
        "manual_flags": {
            491: {
                "file":      "E25_0477_IH_Retro_Elec_Est_Rev_A_AW.xlsx",
                "item_no":   "Item: 2.0",
                "item_name": "Motor Control Centre MCC-PR — BMS interface item",
                "sell_total": 8499.86,
                "mat_cost":   4736.76,
                "mu_A_hrs":   24.0,
                "mu_B_hrs":   16.0,
                "mu_C_hrs":   0.0,
                "reason":     ("MCC-PR has no 512/617 code but interfaces directly with "
                               "Trane BMS controller. Included as BMS per Projections file."),
            },
        },
        "notes": (
            "Validated against Intercontinental_Retrofit_Project_Projections.xlsx. "
            "0477 BMS matches to within €0.50 rounding. "
            "0465 includes Chillers/HTHP items — some overlap with separate Chillers estimate. "
            "Maintenance (€29,592) excluded from revenue scope."
        ),
    },

    # ── ADD NEW PROJECTS BELOW THIS LINE ──────────────────────────────────────
    # Copy and paste the block above, update the fields, add the files to
    # projects/<folder>/ and run python run.py
    #
    # {
    #     "job_number":      "J0XXXX",
    #     "project_name":    "Project Name",
    #     "folder":          "project_folder_name",
    #     "files": [
    #         {
    #             "path":  "estimate_filename.xlsx",
    #             "type":  "EL",      # or "ELM"
    #             "label": "Short label",
    #         },
    #     ],
    #     "project_revenue": {
    #         "estimate_filename.xlsx": 000000.00,
    #     },
    #     "manual_flags": {},
    #     "notes": "",
    # },
]
