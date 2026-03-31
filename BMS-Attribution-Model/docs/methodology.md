# BMS Attribution Model — Methodology & Decision Log
**Panta Contracting Limited — BMS & Controls Department**
Last updated: March 2026

---

## Purpose

This document records every attribution decision made during the development of the BMS Attribution Model. It exists so that anyone running this model in future — including the original author six months later — can understand *why* the rules are what they are, not just *what* they are.

---

## The core problem

BMS works are delivered across mixed-discipline projects (Mechanical, Electrical, BMS, Maintenance). The BMS department currently receives no formal credit in project P&L reporting. This model calculates what percentage of each project's revenue and profit is attributable to BMS, based on BoQ line item analysis.

---

## The estimating template structure

Panta uses two relevant estimating templates:

| Template | File type | Contains BMS? |
|----------|-----------|---------------|
| EL — Electrical | `E25_XXXX_..._Elec_Est_...xlsx` | Yes — primary |
| ELM — Elec-for-Mech | `E25_XXXX_..._Elec4Mech_Est_...xlsx` | Yes — secondary |
| MECH — Mechanical | `M25_XXXX_..._Estimate_...xlsx` | No — out of scope |

Both EL and ELM files share an identical internal structure. The MECH template has a different structure and contains no BMS cost codes.

---

## The ISELL/ICOST cluster rule

### What it is

Every priced item in an EL or ELM estimate is bounded by two markers in column G:
- `ISELL` — opens the block; the Sell Rate and Sell Total are on this row
- `ICOST` — closes the block; total material cost and labour hours echo here

Everything between ISELL and ICOST is a single billable item — one price, one set of costs, one labour allowance.

### Why it was chosen

The cluster is the atomic unit of attribution. Attempting to attribute at individual line-item level (e.g. only the conduit metres that are BMS) would require interpretation of every description. The cluster rule is objective, mechanical, and fully traceable.

### The BMS test

If **any** line item within the cluster has C_CODE **512** or **617**, the entire cluster is BMS-attributed.

- Code **512** = BMS Equipment — hardware, software, controllers, field devices
- Code **617** = Indoor Screened Control Cable — used exclusively for BMS signal wiring

**Decision (March 2026):** All supporting accessories within a BMS cluster (conduit 690, glands 620, junction boxes 570, consumables 875 etc.) are BMS scope. They exist to install the BMS device. The BMS department is entitled to the full revenue and profit for the entire cluster.

### Item identification for the log

- **Item Number**: column D on the ISELL row (e.g. "Item: 10.0")
- **Item Name**: column D on the row immediately below ISELL (e.g. "Flow switch")
- **Source traceability**: source file, sheet name, and Excel row number recorded for every cluster

---

## Cost codes

### BMS anchor codes (trigger attribution)

| Code | Description | Status |
|------|-------------|--------|
| 512 | BMS Equipment | **BMS — primary anchor** |
| 617 | Screened Control Cable | **BMS — co-anchor** |

### Context-dependent codes (BMS when inside a 512/617 cluster)

| Code | Description |
|------|-------------|
| 690 | Conduit (galvanised, flexible, brass adaptors) |
| 620 | Cable glands, cable markers, termination accessories |
| 570 | Junction box, panel enclosure |
| 875 | Cable ties, consumables |
| 705 | Fuse connection units |
| 962 | Crainage and transportation |
| 501 | Erection material, fixings |
| 900 | Labour only — fix/connect |

### Non-BMS codes

| Code | Description | Reason |
|------|-------------|--------|
| 615 | Single-core wiring cable | General electrical, not BMS signal |
| 616 | Multicore power cable (XLPE SWA) | Power distribution, not BMS |
| 650 | Cable tray | General electrical containment |
| 510 | Switchgear, MCBs | General electrical |
| 880 | Design, meetings overhead | Project-wide overhead |
| 901 | Dismantling | Removal of existing plant |
| 902 | Testing & commissioning (electrician) | General electrical T&C |
| 960 | Subcontracting works | General, not BMS-specific |

### Excluded codes

| Code | Description | Reason |
|------|-------------|--------|
| 357 | Electrical for Mechanical (MECH template) | **NOT a BMS signal.** Template placeholder built into mechanical estimating tool. Appears on every pump item. BMS field devices are always booked as 512 in EL/ELM files. |
| RTONLY | Rate only | Zero-value placeholder |
| 1, 2, 3 | Provisional sums | Not committed revenue |
| 821 | Unknown minor accessory | Classified Non-BMS (March 2026). Would only be BMS in unusual circumstances. Cluster rule provides safety net. |

---

## Labour rates

### Grade labels

mu_A, mu_B, mu_C are **pricing rate bands**, not discipline labels. All hours within a BMS cluster are BMS team hours regardless of which rate band was used.

| Grade | Typical description | Cost rate | Sell rate |
|-------|---------------------|-----------|-----------|
| mu_A | Electrical | €27.46/hr | €37.07/hr |
| mu_B | ECA Technician (BMS) | €33.72/hr | €45.54/hr |
| mu_C | ECA Engineer (BMS) | €30.00/hr | €40.50/hr |

### Rate extraction logic

1. Read ML_Sum sheet rows 8–10 for explicit cost rates
2. If row 9 or 10 are blank (as in some EL files), derive: `cost_rate = sell_rate ÷ markup` from Elec_Rates
3. Validate: mu_A derived should match ML_Sum base rate to within €0.05. If it does, derivation is confirmed reliable.

**Validation result (Intercontinental 0477):** mu_A derived = €27.46, ML_Sum = €27.46 ✓. mu_B derived = €33.73 (consistent with 0465 explicit entry of €33.72). mu_C derived = €30.00 ✓.

---

## Manual flags

Some items are legitimately BMS but have no 512/617 code. These must be manually identified and added to `config/settings.py`.

**Intercontinental MCC-PR example:**
- Item: Motor Control Centre MCC-PR, codes 501/570/962
- No 512 or 617 present, so the scanner would exclude it
- However, this panel directly interfaces with the Trane BMS controller
- Decision: include as BMS, document as manual flag
- Value: €8,499.86

---

## Maintenance exclusion

Maintenance contracts (preventative maintenance, 3-year contracts etc.) may have code 512 present (because BMS equipment is being maintained) but are excluded from the project revenue denominator (SUMM PG shows €0 for maintenance sections).

Rule: if item name contains a maintenance keyword AND sell total > €5,000, flag as maintenance and exclude from subtotals. Rows appear in the log and Labour Summary for transparency but are labelled `[MAINT — excl. from subtotal]`.

---

## File scope

| File | Denominator | Notes |
|------|-------------|-------|
| EL estimate | SUMM PG total from that file | Use the pre-tax, pre-discount figure |
| ELM estimate | SUMM PG total from that file | Exclude maintenance section |
| MECH estimate | Out of scope | No 512/617 codes present |

---

## Known limitations and open items

1. **Chillers/HTHP estimate** — The Intercontinental project had a separate Chillers estimate file (M25_0445D_Estimate_Interconti_Chillers_HP...) that was not available. BMS field devices from that file were captured in the Projections file but not by this scanner. Some items appear in both E25-0465 and the Chillers estimate — use Projections file figure (€73,302.16) as authoritative for 0465 BMS until the Chillers file is reconciled.

2. **0477 mu_B cost rate** — ML_Sum for 0477 only has one explicit cost rate. mu_B and mu_C rates are derived from sell_rate/markup. Derivation has been validated against the explicit entries in 0465 and agrees to within €0.01. This is documented in the Attribution Log.
