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
    project_revenue : Total project revenue (EUR, excl VAT).
                      Used as the denominator for BMS Attribution %.
                      If omitted or set to None, auto-extracted from
                      ML_Sum sheet (BOQ TOTAL row, col 4) — works on
                      every EL and ELM file without a SUMM PG sheet.
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
BMS_ANCHOR_CODES = {"512", "617"}   # triggers BMS cluster attribution (EL/ELM project files)
BMS_ML_CODES = {                    # BMS-only standalone estimate files (ML_Sum format)
    "410", "411", "412", "413", "414", "415",
    "416", "417", "418", "419", "420", "421",
}

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
        # Revenue — hardcoded here as these were validated against the Projections file.
        # For new projects leave project_revenue as {} and it will be auto-extracted
        # from ML_Sum (BOQ TOTAL row). Hardcode only if you need to override.
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

    {
        "job_number":      "J01576",
        "project_name":    "Quintano Foods",
        "folder":          "quintano",
        "files": [
            {
                "path":  "E25_0062_Quintano_Foods_Elec_ELV_As_Awarded.xlsx",
                "type":  "EL",
                "label": "E25-0062 (EL)",
            },
            {
                "path":  "ELM25_0085_Quintano_Elec_Mech_HVAC_Est_Rev_B.xlsx",
                "type":  "ELM",
                "label": "E25-0085 (ELM)",
            },
        ],
        # Revenue auto-extracted from ML_Sum — no need to hardcode
        "project_revenue": {},
        "manual_flags": {},
        "notes": "",
    },

    {
        "job_number":      "BMIT",
        "project_name":    "BMIT Project",
        "folder":          "BMIT",
        "files": [
            {
                "path":  "C26_BMIT_ELM_P286.2_R0.1.xlsx",
                "type":  "BMS_ML",
                "label": "C26-BMIT (BMS_ML)",
            },
            {
                "path":  "C26_BMIT_KNX_P286.1_R0.1.xlsx",
                "type":  "BMS_ML",
                "label": "C26-BMIT (BMS_ML)",
            },
        ],
        "project_revenue": {},
        "manual_flags":    {},
        "notes":           "",
    },
    {
        "job_number":      "CortisZebbug",
        "project_name":    "CortisZebbug",
        "folder":          "CortisZebbug",
        "files": [
            {
                "path":  "Cortis Lidl Zebbug_ (SJ SJ554 - R0.1).xlsx",
                "type":  "BMS_TC",
                "label": "Cortis Lidl Zebbug- (SJ SJ554 - R0.1).xlsx (BMS_TC)",
            },
        ],
        "project_revenue": {},
        "manual_flags":    {},
        "notes":           "",
    },
    {
        "job_number":      "CraneCurrencyVSD",
        "project_name":    "CraneCurrencyVSD",
        "folder":          "CraneCurrencyVSD",
        "files": [
            {
                "path":  "Crane Currency - VSD for AHU-09 (SJ572 R0.1).xlsx",
                "type":  "BMS_TC",
                "label": "Crane Curren (BMS_TC)",
            },
        ],
        "project_revenue": {},
        "manual_flags":    {},
        "notes":           "",
    },
    {
        "job_number":      "DomenCarparkPanels",
        "project_name":    "DomenCarparkPanels",
        "folder":          "DomenCarparkPanels",
        "files": [
            {
                "path":  "Dolmen Carpark Panels(SJ574 R0.1).xlsx",
                "type":  "BMS_TC",
                "label": "Dolmen Carpa (BMS_TC)",
            },
        ],
        "project_revenue": {},
        "manual_flags":    {},
        "notes":           "",
    },
    {
        "job_number":      "Getec_Inverter_MDP",
        "project_name":    "Malta Diary",
        "folder":          "Getec_Inverter_MDP",
        "files": [
            {
                "path":  "C26_Getec_Inverter_Panel_SJ586_R0.1_.xlsx",
                "type":  "BMS_ML",
                "label": "C26-Getec (BMS_ML)",
            },
        ],
        "project_revenue": {},
        "manual_flags":    {},
        "notes":           "",
    },
    {
        "job_number":      "LewisPressAHU02",
        "project_name":    "LewisPressAHU02",
        "folder":          "LewisPressAHU02",
        "files": [
            {
                "path":  "Lewis Press (SJ549.1 R1.0).xlsx",
                "type":  "BMS_TC",
                "label": "Lewis Press  (BMS_TC)",
            },
        ],
        "project_revenue": {},
        "manual_flags":    {},
        "notes":           "",
    },
    {
        "job_number":      "Nexent",
        "project_name":    "Multi Tenant Area",
        "folder":          "Nexent",
        "files": [
            {
                "path":  "Nexent_B3L2_SJ583_R0.1.xlsx",
                "type":  "BMS_ML",
                "label": "Nexent-B3L2 (BMS_ML)",
            },
        ],
        "project_revenue": {},
        "manual_flags":    {},
        "notes":           "",
    },
    {
        "job_number":      "NinjaParlour",
        "project_name":    "Ninja Parlour",
        "folder":          "NinjaParlour",
        "files": [
            {
                "path":  "C26_Trident_Ninja_SJ578_R0.3.xlsx",
                "type":  "BMS_ML",
                "label": "C26-Trident (BMS_ML)",
            },
        ],
        "project_revenue": {},
        "manual_flags":    {},
        "notes":           "",
    },
    {
        "job_number":      "PantaHubShowroom_Phase2",
        "project_name":    "Panta Showroom Phase 2",
        "folder":          "PantaHubShowroom_Phase2",
        "files": [
            {
                "path":  "PantaShowroom_Phase2.xlsx",
                "type":  "EL",
                "label": "PantaShowroom-Phase2.xlsx (EL)",
            },
        ],
        "project_revenue": {},
        "manual_flags":    {},
        "notes":           "",
    },
    {
        "job_number":      "PortomasoFCUControllers",
        "project_name":    "PortomasoFCUControllers",
        "folder":          "PortomasoFCUControllers",
        "files": [
            {
                "path":  "Portomaso FCU Controllers (SJ573 R0.2).xlsx",
                "type":  "BMS_TC",
                "label": "Portomaso FC (BMS_TC)",
            },
        ],
        "project_revenue": {},
        "manual_flags":    {},
        "notes":           "",
    },
    {
        "job_number":      "PortomasoSI03_EnergyMeters",
        "project_name":    "Portomaso Phase 3",
        "folder":          "PortomasoSI03_EnergyMeters",
        "files": [
            {
                "path":  "SJ565_Portomaso_SI3_Heat_Meters.xlsx",
                "type":  "BMS_ML",
                "label": "SJ565-Portomaso (BMS_ML)",
            },
        ],
        "project_revenue": {},
        "manual_flags":    {},
        "notes":           "",
    },
    {
        "job_number":      "PortomasoSI04_EnergyMeters",
        "project_name":    "Portomaso Phase 4",
        "folder":          "PortomasoSI04_EnergyMeters",
        "files": [
            {
                "path":  "SJ562_Portomaso_SI4_Heat_Meters.xlsx",
                "type":  "BMS_ML",
                "label": "SJ562-Portomaso (BMS_ML)",
            },
        ],
        "project_revenue": {},
        "manual_flags":    {},
        "notes":           "",
    },
    {
        "job_number":      "Portomaso_L4_Dampers",
        "project_name":    "Portomaso L4 Dampers",
        "folder":          "Portomaso_L4_Dampers",
        "files": [
            {
                "path":  "C26_Portomaso_L4_Dampers_SJ585.1_R0.1.xlsx",
                "type":  "BMS_ML",
                "label": "C26-Portomaso (BMS_ML)",
            },
        ],
        "project_revenue": {},
        "manual_flags":    {},
        "notes":           "",
    },
    {
        "job_number":      "STMicroExtractFans",
        "project_name":    "STMicroExtractFans",
        "folder":          "STMicroExtractFans",
        "files": [
            {
                "path":  "ST MicroElectronics KK2 (SJ564 R0.1).xlsx",
                "type":  "BMS_TC",
                "label": "ST MicroElec (BMS_TC)",
            },
        ],
        "project_revenue": {},
        "manual_flags":    {},
        "notes":           "",
    },
    {
        "job_number":      "STPropertiesPumpInverter",
        "project_name":    "STPropertiesPumpInverter",
        "folder":          "STPropertiesPumpInverter",
        "files": [
            {
                "path":  "ST Properties Borehole Pump Inverter Replacement(SJ577 R0.1).xlsx",
                "type":  "BMS_TC",
                "label": "ST Propertie (BMS_TC)",
            },
        ],
        "project_revenue": {},
        "manual_flags":    {},
        "notes":           "",
    },
    {
        "job_number":      "ST_Micro_Exhaust_Fans",
        "project_name":    "KK2 General Exhaust Replacement",
        "folder":          "ST_Micro_Exhaust_Fans",
        "files": [
            {
                "path":  "E25_0559_ST_Electronics_Exhaust_Replacement_Est_.xlsx",
                "type":  "EL",
                "label": "E25-0559 (EL)",
            },
            {
                "path":  "ELM25_0560_ST_Electronics_Exhaust_Replacement_Est_.xlsx",
                "type":  "ELM",
                "label": "ELM25-0560 (ELM)",
            },
        ],
        "project_revenue": {},
        "manual_flags":    {},
        "notes":           "",
    },
    {
        "job_number":      "Sunny_Coast",
        "project_name":    "Sunny Cost and Lido Redevelopment",
        "folder":          "Sunny_Coast",
        "files": [
            {
                "path":  "EL25_0700_Mock_Up_Room_Elec_Estim.xlsx",
                "type":  "EL",
                "label": "EL25-0700 (EL)",
            },
            {
                "path":  "ELM25_0699_Mock_Up_Room_Elec_for_Mech.xlsx",
                "type":  "ELM",
                "label": "ELM25-0699 (ELM)",
            },
        ],
        "project_revenue": {},
        "manual_flags":    {},
        "notes":           "",
    },
    {
        "job_number":      "TridentParkChillerCabling",
        "project_name":    "TridentParkChillerCabling",
        "folder":          "TridentParkChillerCabling",
        "files": [
            {
                "path":  "Trident Park Chiller Cabling (SJ62.1 R0.1).xlsx",
                "type":  "BMS_TC",
                "label": "Trident Park (BMS_TC)",
            },
        ],
        "project_revenue": {},
        "manual_flags":    {},
        "notes":           "",
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