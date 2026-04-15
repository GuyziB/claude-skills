# BMS Attribution Model — Claude Code Instructions
Panta Contracting Limited

## What this project does
Scans Panta Contracting electrical estimate files for BMS-attributed work.
Produces an Excel output per project with attribution log, labour hours and P&L.

Handles three estimate formats:
- **EL / ELM** — full project estimates embedded in a larger Panta job (ISELL/ICOST cluster rule)
- **BMS_ML** — standalone BMS estimate in the same ML_Sum/Elec_Est format, but entire file is BMS
- **BMS_TC** — standalone BMS panel/small-job estimate with TotalCosts (Hrs) sheet

## How to run
```
python run.py              # all registered projects
python run.py J01576       # one project by job number
```

## Adding a new project — zero-config workflow
1. Create `projects/J05226_ST_Micro_Exhaust_Fans/`  (J0XXXX_ prefix + human-readable name)
2. Drop estimate xlsx files into it
3. Run: `python run.py J05226`  (job number only — discovers folder by prefix)
   - File type auto-detected from sheet names (see below)
   - Project name auto-read from ML_Sum or TotalCosts sheet
   - Entry written to config/settings.py automatically
   - Attribution runs immediately

## File type detection (automatic — no manual setting needed)
`core/discover.py` inspects sheet names when a new folder is registered:

| Detected type | Condition |
|---|---|
| `BMS_TC` | Has `TotalCosts (Hrs)` sheet |
| `BMS_ML` | Has `ML_Sum` + `Elec_Est` + any cost code 410–421 in Elec_Est |
| `ELM`    | Has `ML_Sum`, filename starts with ELM |
| `EL`     | Has `ML_Sum`, everything else |

## Attribution rules — never change these

### EL / ELM (project estimates)
- Scan `Elec_Est` sheet for ISELL/ICOST blocks
- If ANY line item in the block has C_CODE **512 or 617** → entire block is BMS
- DO NOT add code 570 as a BMS anchor — it overcounts
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
- Grades tracked: mu_A, mu_B, mu_C, mu_D

### BMS_ML (standalone BMS estimates with ML_Sum)
- BMS cost codes: 410 Controllers, 411 Controller Modules, 412 Signal Cable,
  413 Modbus Cable, 414 CAT Cable, 415 Control Panels, 416 Field Device,
  417 DALI/KNX, 418 Fire Equipment, 419 Data Networks, 420 Panel Building,
  421 Terminations
- If ANY of these codes appear → entire file is 100% BMS
- All ISELL/ICOST clusters included (force_bms=True), no per-cluster code check
- Revenue extracted from ML_Sum (same method as EL/ELM)

### BMS_TC (standalone BMS panel/small-job estimates)
- Sheet: `TotalCosts (Hrs)` — no ISELL/ICOST structure
- Read single summary row: `TOTAL (Excl. VAT)` in col 3
- Revenue  = col 43 (TOTAL with mark-up)
- Material = col 4  (material cost excl. VAT)
- Hours    = col 34 (Tradesman pair hrs) + col 35 (Technician pair hrs) + col 36 (Engineer man hrs)
- Labour cost uses DEFAULT_RATES (no rate sheet in these files)
- Entire file is 100% BMS by definition

## Revenue auto-extraction (EL / ELM / BMS_ML)
ML_Sum sheet → find `Discy.A-J` header row → sum each grade row below using
col 4 (discounted selling) if populated, else col 5 as fallback → stop at `Discy.L`.
Works for both EL (col 4 always populated) and ELM (some grades only have col 5).
SUMM PG sheet is NOT needed.

## Confirmed results (do not regress)
| Project | Revenue | BMS | % | Hrs | GP |
|---|---|---|---|---|---|
| J01577 Intercontinental | €551,834 | €271,780 | 49.3% | 3,032 | €67,044 |
| J01576 Quintano EL | €1,478,219 | €9,982 | 0.7% | 90 | 18.8% margin |
| J01576 Quintano ELM | €80,021 | €41,122 | 51.4% | 375 | 0.0% (correct — no labour markup) |

## Project file locations
| File | Purpose |
|---|---|
| `config/settings.py` | Project registry — auto-updated by discover workflow |
| `core/scanner.py` | ISELL/ICOST cluster scan + BMS_TC scanner |
| `core/rates.py` | Labour rate extraction + ML_Sum revenue extraction |
| `core/attribution.py` | Orchestrates scanning, routes by file type |
| `core/discover.py` | Auto-discovery: detects file type, registers in settings.py |
| `outputs/excel_builder.py` | Builds Excel output workbook |
| `projects/intercontinental/` | Intercontinental estimate files |
| `projects/quintano/` | Quintano estimate files |

## Session management
Be token-efficient. Before writing any code, state your plan in one sentence.
Make targeted edits only — do not rewrite files unnecessarily.
After completing a task, update this file with what was done, then stop.
