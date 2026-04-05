# BMS Attribution Model — Claude Code Instructions
Panta Contracting Limited

## What this project does
Scans Panta Contracting electrical estimate files (EL and ELM format) for 
BMS-attributed work using the ISELL/ICOST cluster rule. Produces an Excel 
output with attribution log, labour hours and P&L per project.

## How to run
python run.py              # all projects
python run.py J01576       # one specific project by job number

## Attribution rule — never change this
- Scan Elec_Est sheet for ISELL/ICOST blocks
- If ANY line item in the block has C_CODE 512 or 617 → entire block is BMS (EL/ELM)
- DO NOT add code 570 as a BMS anchor — it overcounts

## Standalone BMS estimate file types
Two additional file types for standalone BMS team jobs (not integrated into bigger projects):

**BMS_ML** — has ML_Sum + Elec_Est, uses BMS cost codes 410–421.
  If ANY code 410–421 appears in the file → entire file is 100% BMS.
  All ISELL/ICOST clusters included regardless of per-cluster codes.
  Revenue extracted from ML_Sum as normal.

**BMS_TC** — has TotalCosts (Hrs) sheet (panel/small-job estimator format).
  No ISELL/ICOST structure. Single summary row 'TOTAL (Excl. VAT)'.
  Revenue = col 43 (TOTAL with mark-up). Material = col 29. Hours = cols 34+35+36.
  Entire file is 100% BMS by definition. Labour cost uses DEFAULT_RATES.

File type is auto-detected by discover.py from sheet names — no manual setting needed.
- All hours within a BMS cluster are BMS team hours regardless of grade
- Item Number = col D on ISELL row, Item Name = col D one row below
- Sell Total = col F (index 5) on ISELL row
- ISELL/ICOST marker = col G (index 6)
- C_CODE = col H (index 7)
- Grade label on d> SUBI rows = col K (index 10)
- Labour hours on d> SUBI rows = col T (index 19)
- Total Material Cost on ISELL row = col AM (index 38)
- Total Labour Hours on ISELL row = col AN (index 39)
- SUM all d> rows per grade — do not overwrite, accumulate with +=

## Current task — J01576 Quintano Foods

### What was fixed this session
- run.py: added `sys.stdout.reconfigure(encoding="utf-8")` to fix cp1252 crash on Windows
- core/scanner.py: added mu_D to grade_hrs dict and lab_cost loop — fixes hour check
  warning on Item 4.028 (sum was missing mu_D hours)
- core/rates.py: rewrote get_project_revenue() to sum grade rows using
  col 4 if available, else col 5 as fallback — fixes ELM revenue extraction
  (was returning €32,714 instead of €80,021)

### Confirmed results
J01577 Intercontinental: revenue €551,834 | BMS €271,780 (49.3%) | 3,032 hrs | GP €67,044 ✓
J01576 Quintano EL:      revenue €1,478,219 | BMS €9,982 (0.7%) | 90 hrs | margin 18.8%
J01576 Quintano ELM:     revenue €80,021 | BMS €41,122 (51.4%) | 375 hrs | margin 0.0% (correct — no labour markup in file)

### ELM 0% gross margin — explained
ELM file estimated at labour cost = sell rate (both 27.46/hr). sell_total = mat_cost + lab_cost
exactly for every cluster — no labour profit is built in at item level. This is correct.

### How revenue auto-extraction works
ML_Sum sheet → find 'Discy.A-J' header row → sum each grade row below it using
col 4 (discounted selling) if populated, else col 5 as fallback → stop at Discy.L.
Works for both EL (single/multi-grade with col 4 always populated) and ELM
(some grades only have col 5). SUMM PG sheet is NOT needed.

## Adding a new project — zero-config workflow
1. Create projects/J05226_ST_Micro_Exhaust_Fans/  (J0XXXX_ prefix + human-readable name)
2. Drop estimate xlsx files into it (ELM* files auto-detected as ELM type, E* as EL)
3. Run: python run.py J05226   (job number only — discovers the folder by prefix)
   — Project name auto-read from ML_Sum sheet
   — Entry written to config/settings.py automatically (stores full folder name)
   — Attribution runs immediately

## Project file locations
config/settings.py        — project registry (auto-updated by discover workflow)
core/scanner.py           — ISELL/ICOST cluster scan engine
core/rates.py             — labour rate extraction + revenue auto-extraction
core/attribution.py       — orchestrates scanning per project
core/discover.py          — auto-discovery: scans new folders, registers in settings.py
outputs/excel_builder.py  — builds Excel output
projects/quintano/        — Quintano estimate files live here
projects/intercontinental/— Intercontinental estimate files

## Session management
Be token-efficient. Before writing any code, state your plan in one 
sentence. Make targeted edits only — do not rewrite files unnecessarily. 
After completing the task update the Current task section of this file 
with what was done and what is next, then stop.
```

And when you prompt it, say:
```
Read CLAUDE.md and complete the current task. Be concise.
