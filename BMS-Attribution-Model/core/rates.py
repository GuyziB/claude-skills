"""
core/rates.py
=============
Extracts labour grade descriptions, sell rates, and cost rates from
an estimate file.

Rate extraction priority:
    1. ML_Sum sheet rows 8-10 — explicit cost rates (present in ELM files)
    2. Derived: sell_rate / markup from Elec_Rates sheet
       (used when ML_Sum only has the base site rate)

Validation: mu_A derived cost should match ML_Sum base rate to within €0.05.
If it does, derivation is confirmed reliable for mu_B and mu_C.
"""

import pandas as pd
from config.settings import DEFAULT_RATES


def get_rates(filepath: str) -> dict:
    """
    Extract grade descriptions, sell rates, and cost rates from an estimate file.

    Returns a dict keyed by grade name:
    {
        "mu_A": {
            "desc":       "Electrical",
            "sell_rate":  37.07,
            "cost_rate":  27.46,
            "source":     "ML_Sum (matches derived)"
        },
        "mu_B": { ... },
        "mu_C": { ... },
        "mu_D": { ... },
    }
    """
    rates = {}

    try:
        df_er = pd.read_excel(filepath, sheet_name="Elec_Rates", header=None)
        df_ml = pd.read_excel(filepath, sheet_name="ML_Sum",     header=None)
    except Exception as e:
        print(f"  [rates] WARNING: Could not read rate sheets from {filepath}: {e}")
        return {g: {**v, "sell_rate": 0.0, "source": "default"} 
                for g, v in DEFAULT_RATES.items()}

    # ── Step 1: Read ML_Sum explicit cost rates (rows 8, 9, 10) ──────────────
    ml_cost = {}
    for i, grade in enumerate(["mu_A", "mu_B", "mu_C"], 8):
        try:
            v = df_ml.iloc[i, 3]
            if pd.notna(v) and float(v) > 0:
                ml_cost[grade] = float(v)
        except (IndexError, ValueError):
            pass

    # ── Step 2: Read Elec_Rates — descriptions, markups, sell rates ──────────
    for i, grade in enumerate(["mu_A", "mu_B", "mu_C", "mu_D"], 1):
        try:
            desc      = str(df_er.iloc[i, 5]).strip()
            markup_v  = df_er.iloc[i, 18]
            sell_v    = df_er.iloc[i, 19]
            markup    = float(markup_v) if pd.notna(markup_v) else 1.0
            sell_rate = float(sell_v)   if pd.notna(sell_v)   else 0.0
        except (IndexError, ValueError):
            desc      = DEFAULT_RATES.get(grade, {}).get("desc", grade)
            markup    = 1.0
            sell_rate = 0.0

        # Derive cost rate from sell rate and markup
        derived = round(sell_rate / markup, 2) if markup > 0 else 0.0

        # Determine final cost rate and source
        if grade in ml_cost:
            if abs(ml_cost[grade] - derived) < 0.05:
                cost_rate = ml_cost[grade]
                source    = "ML_Sum (matches derived)"
            else:
                cost_rate = ml_cost[grade]
                source    = "ML_Sum (explicit)"
        else:
            cost_rate = derived
            source    = "Derived (sell÷markup)"

        rates[grade] = {
            "desc":      desc,
            "sell_rate": sell_rate,
            "cost_rate": cost_rate,
            "source":    source,
        }

    return rates


def format_rate_legend(rates: dict) -> str:
    """Return a one-line rate summary string for display in Excel."""
    parts = []
    for grade in ["mu_A", "mu_B", "mu_C"]:
        if grade in rates:
            r = rates[grade]
            parts.append(
                f"{grade} = {r['desc']} "
                f"(cost €{r['cost_rate']:.2f}/hr, "
                f"sell €{r['sell_rate']:.2f}/hr "
                f"[{r['source']}])"
            )
    return "  |  ".join(parts)
