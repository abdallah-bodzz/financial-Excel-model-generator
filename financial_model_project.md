# Financial Model Generator — Project Documentation

## What This Is

A Python script (`generate_model_v4.py`) that produces a fully Excel-driven, 3-statement financial model as an `.xlsx` file. Python only builds structure, formulas, and styling. Excel does every calculation. No numbers are computed in Python — they all live as Excel formulas that recalculate when assumptions change.

---

## Project Structure

```
financial_model/
├── config.json           ← all numeric inputs (edit this to change the model)
├── generate_model.py     ← the generator script
└── output/
    └── Financial_Model.xlsx   ← generated file (open in Excel)
```

---

## config.json Reference

```json
{
  "forecast_years": 5,
  "base_revenue": 1000000,
  "initial_cash": 150000,
  "initial_ppe": 500000,
  "initial_debt": 200000,
  "scenarios": {
    "Base":  { "revenue_growth": 0.10, "cogs_pct": 0.45 },
    "Best":  { "revenue_growth": 0.18, "cogs_pct": 0.40 },
    "Worst": { "revenue_growth": 0.04, "cogs_pct": 0.55 }
  },
  "assumptions": {
    "rd_pct": 0.12,
    "sga_pct": 0.15,
    "tax_rate": 0.21,
    "depreciation_rate": 0.10,
    "capex_pct": 0.05,
    "dso_days": 45,
    "dio_days": 60,
    "dpo_days": 30,
    "interest_rate": 0.05,
    "principal_repayment": 0,
    "wacc": 0.09,
    "perpetual_growth": 0.025
  }
}
```

**Rules:**
- `forecast_years` drives the number of columns on every sheet. Change it and the whole model resizes.
- Add scenarios freely — the dropdown and INDEX/MATCH logic are range-based and scale automatically.
- `principal_repayment` = 0 means bullet debt (repaid at maturity). Set to a positive number for amortizing.
- `perpetual_growth` must be < `wacc` or the Terminal Value formula will return an error (`WACC ≤ g`).

---

## Sheet Map

| Sheet | Purpose | Inputs live here? |
|---|---|---|
| Start Here | Quick-start guide, color legend, sensitivity instructions | No |
| Control Panel | All assumptions, scenario table, scenario selector dropdown | **Yes** |
| P&L | Income statement: Revenue → Net Income | Year 1 Revenue only |
| Balance Sheet | Assets = L+E every year, balance check row | Opening Cash, PPE, Debt |
| Cash Flow | Indirect method, reconciles NI to ending cash | No |
| Valuation | FCFF DCF → EV → Equity Value, implied multiples | No |
| Sensitivity | Data Table shell: Revenue Growth × COGS % → NI Year N | Growth/COGS headers |
| Insights | KPI summary: margins, CAGR, break-even, checks | No |
| Dashboard | 3 charts: Revenue, Net Income, EBITDA Margin | No |

---

## How the Model Links Together

```
Control Panel assumptions
        │
        ▼
    P&L (Income Statement)
    Revenue Y1 (hardcoded) → Y2..N via growth rate from Control Panel
    COGS = Revenue × COGS% from Control Panel
    R&D, SG&A = Revenue × rates from Control Panel
    Depreciation = opening PP&E × rate (opening PP&E from prior BS year)
    Interest = opening Debt × rate (opening Debt from prior BS year)
    Taxes = MAX(EBT × rate, 0)   ← no negative taxes
        │
        ▼
    Balance Sheet
    Cash Y1 = hardcoded
    Cash YN = Cash Y(N-1) + NChg YN   ← uses CURRENT year CF, not prior year
    AR, Inventory, AP = average-period × days/365
    PP&E YN = PP&E Y(N-1) - CF_Capex_YN - Depr_YN
    Debt YN = MAX(0, Debt Y(N-1) - repayment)   ← no negative debt
    RE Y1 = NI Y1 (not 0!)
    RE YN = RE Y(N-1) + NI YN
    Equity Plug Y1 = TA - AP - Debt - RE   (absorbs opening capital)
    Equity Plug Y2+ = constant = Y1 plug   (paid-in capital doesn't change)
        │
        ▼
    Cash Flow
    Net Income + Depr + ΔWC + Capex → Net Change in Cash
    Feeds back into Balance Sheet Cash (above)
        │
        ▼
    Valuation
    FCFF = EBIT×(1-t) + D&A + ΔWC + Capex
    EV = NPV(FCFF) + PV(Terminal Value)
    Equity Value = EV - Net Debt
```

---

## Key Design Decisions & Why

### No named ranges anywhere
Named ranges in openpyxl fail silently on many Excel versions → every formula that would use a named range uses a direct `'Sheet'!$B$N` reference instead. This is the root cause of `#NAME?` errors in most Python-generated models.

### Equity Plug treatment
The BS has no "opening equity" input, so Year 1 would never balance without a plug. The plug = `TA - AP - Debt - RE` at Year 1 only. After that it's constant. This correctly represents paid-in capital that was used to fund the opening assets. Do not recompute it every year — that would make equity circular.

### Cash formula direction
`Cash YN = Cash Y(N-1) + NChg YN` — uses the **same column** as the BS year. The off-by-one mistake (`col-1`) is the most common balance-sheet-breaking bug in Python-generated models. The CF sheet's `Net Change in Cash` col B = Year 1, col C = Year 2, etc. The BS Cash formula for Year 2 (col C) must reference CF col C.

### Interest uses opening debt only
Average-debt interest (`(open + close)/2 × rate`) creates a circular chain: Interest → NI → CF → Debt → Interest. Excel can resolve this with iterative calculation enabled, but most users don't have that on. Opening-debt interest is the industry-standard safe approach for non-circular models.

### Average-period working capital
`AR = ((Revenue_curr + Revenue_prev) / 2) × DSO / 365` instead of just `Revenue_curr × DSO / 365`. This reflects that the balance sheet is a snapshot at period end, not a point-in-time reading of a single day's revenue.

### Data Table for Sensitivity (not hardcoded formulas)
A Data Table is the correct Excel-native approach. It means every cell in the sensitivity matrix is computed by Excel using the actual live model — so if you change the scenario or any assumption, the whole table updates. Hardcoded LET formulas in each cell look similar but are disconnected from the model.

### Taxes: MAX(EBT × rate, 0)
Real companies don't receive tax refunds in loss years (in most jurisdictions, at least not immediately). Using `MAX(...)` prevents negative tax expense which would inflate Net Income on loss scenarios.

---

## Sensitivity Sheet — How to Activate

The script creates the Data Table shell but cannot auto-run the Data Table (Excel doesn't expose this via file format). After opening:

1. Go to **Sensitivity** sheet
2. Select `B3:H10` (the full body including corner, headers, and body cells)
3. **Data** → **What-If Analysis** → **Data Table**
4. Row input cell: `'Control Panel'!$B$24` (Effective Revenue Growth)
5. Column input cell: `'Control Panel'!$B$25` (Effective COGS %)
6. Click OK

The exact cell addresses are printed in the terminal when you run the script and shown in the orange instruction banner on the sheet.

---

## Checks Built Into the Model

| Check | Location | What it tests |
|---|---|---|
| Balance Check | Balance Sheet row 14 | TA − L&E = 0 every year. Green = correct, red = error |
| Model Health | Control Panel | SUM of all balance check cells |
| Sources & Uses | Insights | Ending Cash = Opening Cash + cumulative NChg Y2..YN |
| WACC > g | Valuation, Control Panel flag | Terminal Value is only valid when WACC > perpetual growth |
| Negative cash | Insights | Flags if ending cash < 0 |
| Balance Check Status | Insights | Confirms ABS(balance check) < 1 for rounding tolerance |

---

## Common Errors & Fixes

| Error | Cause | Fix |
|---|---|---|
| `#NAME?` everywhere | Named ranges didn't register | Use `'Sheet'!$B$N` direct references only |
| Balance check ≠ 0 from Year 2 | Cash formula uses `col-1` CF instead of `col` | `Cash YN = Cash Y(N-1) + NChg YN` (same column) |
| RE grows faster than assets | RE Y1 hardcoded to 0 instead of NI Y1 | `RE Y1 = NI Y1`, then `RE YN = RE Y(N-1) + NI YN` |
| Debt goes negative | No floor on repayment | `MAX(0, prev_debt - repayment)` |
| Terminal Value = error | `perpetual_growth >= wacc` | Wrap in `IFERROR`, add flag in Control Panel |
| Circular reference warning | Average-debt interest | Use opening debt only |
| Protection blocks edits | Inputs not unlocked before `ws.protection.sheet = True` | Set `Protection(locked=False)` on input cells explicitly |
| Data Table doesn't populate | Script can't trigger it — manual step required | See instructions above |

---

## Extending the Model

**Add a new assumption:** Add an entry to `assumption_order` in the script. The row number auto-increments. All sheets reference Control Panel by row number via `cp_ref(key)`, so they pick it up automatically.

**Add a new scenario:** Add to `config.json` `scenarios` dict. The scenario table, dropdown validation range, and INDEX/MATCH formulas all use `len(SCENARIOS)` — they extend automatically.

**Add a new P&L line:** Add a key to `pl_rows` (increment rows if needed), add a label in `pl_label_map`, write the formula in the `for col in range(2, YEARS+2)` loop.

**Change forecast horizon:** Edit `forecast_years` in `config.json`. Every loop uses `YEARS`, every range reference uses `YEARS+1` or `YEARS+2` as bounds. Charts auto-extend.

**Add a new KPI to Insights:** Append a tuple to the `kpis` list: `(label, formula_string, fmt, comment_formula)`. `fmt` is `"percent"`, `"money"`, or `"text"`.