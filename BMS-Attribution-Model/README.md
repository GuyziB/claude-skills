# BMS Attribution Model
**Panta Contracting Limited — BMS & Controls Department**

Automatically extracts BMS revenue, labour hours, and gross profit from Panta estimating files (EL and ELM formats) using the ISELL/ICOST cluster attribution rule.

---

## What it does

Scans electrical estimate files (`.xlsx`) for BMS-attributed work using cost codes **512** and **617** as anchor signals within ISELL/ICOST blocks. Produces a formatted Excel report with:

- Full attribution log (one row per BMS cluster, with source file and Excel row reference)
- Labour hours split by grade (mu_A / mu_B / mu_C)
- Gross profit per cluster and per project
- Combined P&L summary — all formula-driven

---

## Folder structure

```
BMS-Attribution-Model/
├── run.py                      ← Entry point — run this
├── requirements.txt            ← Python dependencies
├── README.md                   ← This file
│
├── config/
│   └── settings.py             ← Add new projects here
│
├── core/
│   ├── scanner.py              ← ISELL/ICOST cluster scan engine
│   ├── rates.py                ← Labour rate extraction
│   └── attribution.py          ← BMS attribution logic
│
├── outputs/
│   └── excel_builder.py        ← Builds the output Excel workbook
│
├── projects/
│   └── intercontinental/       ← One subfolder per project
│       ├── [estimate files]
│       └── [output xlsx]
│
└── docs/
    └── methodology.md          ← Full methodology and decision log
```

---

## How to run

### First time setup
```powershell
cd BMS-Attribution-Model
pip install -r requirements.txt
```

### Running the model
```powershell
python run.py
```

This will process all projects defined in `config/settings.py` and produce output Excel files in each project folder.

### Adding a new project
1. Create a subfolder under `projects/` with the job number or name
2. Copy the estimate `.xlsx` files into that subfolder
3. Open `config/settings.py` and add the project definition (follow the existing example)
4. Run `python run.py`

---

## Attribution rule (summary)

- Scan every `Elec_Est` sheet for ISELL/ICOST blocks
- If **any** line item within the block has cost code **512** or **617** → entire block is BMS
- Extract: sell total, material cost, total labour hours, grade split (mu_A/B/C), labour cost, gross profit
- Log the source file, sheet name, and Excel row number for every attributed block
- Applies to **EL** (Electrical) and **ELM** (Elec-for-Mech) estimate files only
- Mechanical-only files (MECH template) contain no 512/617 codes and are out of scope

## Cost code reference

| Code | Description | BMS Attribution |
|------|-------------|-----------------|
| 512  | BMS Equipment | **100% BMS — primary anchor** |
| 617  | Screened Control Cable | **100% BMS — co-anchor** |
| 690  | Conduit | BMS when in 512/617 cluster |
| 620  | Cable Glands / Markers | BMS when in 512/617 cluster |
| 570  | Junction Box / Panel | BMS when in 512/617 cluster |
| 616  | Power Cable (XLPE SWA) | Non-BMS |
| 615  | Single-core Wiring Cable | Non-BMS |
| 357  | Elec-for-Mech placeholder | **Not a BMS signal — exclude** |

## Labour rate logic

Rates extracted per file in this priority order:
1. ML_Sum sheet rows 8–10 (explicit cost rates — present in ELM files)
2. Derived: `sell_rate ÷ markup` from Elec_Rates sheet (used when ML_Sum only has base rate)

| Grade | Description | Typical cost rate |
|-------|-------------|-------------------|
| mu_A  | Electrical  | €27.46/hr |
| mu_B  | ECA Technician (BMS) | €33.72/hr |
| mu_C  | ECA Engineer (BMS)   | €30.00/hr |

**All hours within a BMS cluster are BMS team hours regardless of rate band.**

---

## Version history

| Version | Date | Description |
|---------|------|-------------|
| 3.1 | March 2026 | Formula-driven subtotals and combined P&L |
| 3.0 | March 2026 | Labour hours and P&L added |
| 2.0 | March 2026 | Stage 2 — BoQ extraction and attribution log |
| 1.0 | March 2026 | Stage 1 — Cost code mapping table |
