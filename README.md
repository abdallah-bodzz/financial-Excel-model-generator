# Financial Model Generator

## Overview

A Python script that produces a fully Excel‑driven, 3‑statement financial model as an `.xlsx` file.  
Python only builds the structure, formulas, and styling. Excel performs every calculation – all numbers live as formulas that recalculate when assumptions change.

## Project Structure

```
financial_model/
├── config.json           # All numeric inputs (edit this to change the model)
├── generate_model.py     # The generator script
└── output/
    └── Financial_Model.xlsx   # Generated file (open in Excel)
```

## Configuration (`config.json`)

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

### Rules

- `forecast_years` determines the number of forecast columns on every sheet.
- Add or remove scenarios freely – the dropdown and `INDEX/MATCH` logic scale automatically.
- `principal_repayment = 0` means bullet debt (repaid at maturity). Set a positive value for amortizing debt.
- `perpetual_growth` must be less than `wacc`; otherwise the terminal value formula returns an error.

## Sheets in the Generated Model

| Sheet | Purpose |
|-------|---------|
| Start Here | Quick‑start guide, colour legend, sensitivity instructions |
| Control Panel | All assumptions, scenario table, scenario selector dropdown |
| P&L | Income statement: Revenue → Net Income |
| Balance Sheet | Assets = Liabilities + Equity every year, with a balance check row |
| Cash Flow | Indirect method, reconciles Net Income to ending cash |
| Valuation | Unlevered DCF (FCFF) → Enterprise Value → Equity Value, plus implied multiples |
| Sensitivity | Data Table shell for stress‑testing Year N Net Income |
| Insights | KPI summary: margins, CAGR, break‑even year, cash runway, model checks |
| Dashboard | Three charts: Revenue, Net Income, EBITDA Margin |

## Key Design Decisions

### No Named Ranges
Named ranges in `openpyxl` fail silently on many Excel versions. Every formula uses direct `'Sheet'!$B$N` references.

### Equity Plug
The Balance Sheet has no "opening equity" input. Year 1 equity plug is calculated as `Total Assets − AP − Debt − Retained Earnings`. This value is then held constant for all future years (paid‑in capital does not change).

### Cash Formula Direction
`Cash Yn = Cash Y(n-1) + Net Change in Cash Yn` – uses the **same column** from the Cash Flow sheet. This avoids the common off‑by‑one error.

### Interest Uses Opening Debt Only
Average‑debt interest creates a circular reference. Opening‑debt interest is the safe, non‑circular approach.

### Average‑Period Working Capital
`AR = ((Revenue_curr + Revenue_prev)/2) × DSO / 365` – reflects that the balance sheet is a snapshot at period end.

### Taxes: `MAX(EBT × rate, 0)`
Prevents negative tax expense in loss years (no immediate tax refunds assumed).

## Activating the Sensitivity Data Table

The script creates the Data Table shell but cannot auto‑run it (Excel does not expose this via file format). After opening the generated file:

1. Go to the **Sensitivity** sheet.
2. Select the full matrix (e.g., `B3:H10`).
3. **Data** → **What‑If Analysis** → **Data Table**.
4. Row input cell: `'Control Panel'!$B$24` (Effective Revenue Growth).
5. Column input cell: `'Control Panel'!$B$25` (Effective COGS %).
6. Click **OK**.

The exact cell addresses are printed in the terminal and shown in an orange instruction banner on the sheet.

## Built‑in Model Checks

| Check | Location | Description |
|-------|----------|-------------|
| Balance Check | Balance Sheet row 14 | `Total Assets − Total L&E = 0` every year. Green = correct, red = error. |
| Model Health | Control Panel | Sum of all balance check cells. |
| Sources & Uses | Insights | Verifies that `Ending Cash = Opening Cash + sum(Net Change in Cash Y2..YN)`. |
| WACC > g | Valuation & Control Panel | Flags when `perpetual_growth ≥ WACC` (invalid terminal value). |
| Negative Cash | Insights | Flags if ending cash is negative. |

## Extending the Model

- **Add an assumption**: Add an entry to `assumption_order` in the script. The row number auto‑increments; all sheets reference it via `cp_ref()`.
- **Add a scenario**: Add a new key to `scenarios` in `config.json`. The scenario table, dropdown, and `INDEX/MATCH` formulas extend automatically.
- **Change forecast horizon**: Edit `forecast_years` in `config.json`. All loops and column ranges adjust automatically. Charts also auto‑extend.
- **Add a KPI**: Append a tuple to the `kpis` list in the script: `(label, formula, "percent"|"money"|"text", comment_formula)`.

## Requirements

- Python 3.8+
- `openpyxl` library

Install with:
```bash
pip install openpyxl
```

## Usage

```bash
python generate_model.py
```

The generated Excel file will be placed in the `output/` directory.
