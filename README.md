# Headcount and staffing forecast

Plan, loaded-cost rollforward, and reforecast for a fictional CPG company (**Northline Consumer Products**): ~280–310 FTE across ops, supply chain, commercial, and G&A.

**File to open:** `Northline_Headcount_Staffing_Forecast.xlsx`

## What you will see

- Department drivers: opening FTE, fully loaded cost per FTE, attrition, merit, hire lead time
- Monthly headcount plan: opening + hires − exits = ending FTE (mix of inputs and formulas)
- Loaded-cost rollforward with a April merit step-up; plan vs forecast $ variance
- Latest view vs original plan (YTD actuals through an as-of month, then forecast)
- A one-pager with FTE, cost, pipeline, and attrition

Yellow cells with blue font are inputs. Black font is formulas.

## How to use

1. Open `01_Assumptions` and change the yellow cells (loaded cost, attrition, raise, as-of month).
2. Edit plan hires and other exits on `02_Headcount_Plan`. Opening, attrition exits, and ending FTE are formulas.
3. Read the cost rollforward and plan vs forecast variance on `03_Cost_Bridge`.
4. Edit latest-view hires / exits on `04_Reforecast` and compare to the original plan.
5. Use `05_Dashboard` as the one-pager.

## Tabs

| Tab | Role |
| --- | --- |
| `00_Cover` | Purpose and how to use |
| `01_Assumptions` | Departments, loaded cost, attrition, raise, as-of month |
| `02_Headcount_Plan` | Opening, hires, exits, ending FTE by dept × month |
| `03_Cost_Bridge` | Monthly loaded cost; plan vs forecast variance |
| `04_Reforecast` | Latest view vs original plan |
| `05_Dashboard` | One-pager: FTE, cost, pipeline, attrition |
| `06_Data_Dictionary` | Field definitions |

## Stack

Excel (formulas only — no VBA). Built so another analyst can inherit the file from the data dictionary. Cost in $000s.

## Not included on purpose

- Live HRIS feeds or named-employee files
- Confidential employer data
- Benefits detail below fully loaded cost (this is the *staffing and $ run-rate* view)

All sample numbers are fictional.

[Profile](https://github.com/saisiri1207) · [Portfolio](https://saisiri1207.github.io) · [LinkedIn](https://www.linkedin.com/in/saisiri1207) · [bandarusaisiri1207@gmail.com](mailto:bandarusaisiri1207@gmail.com)
