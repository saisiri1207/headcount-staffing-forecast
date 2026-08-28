# Headcount and loaded-cost reforecast

Method note for a 12-month staffing model. Built for the same fictional company used across this profile (**Northline Consumer Products**).

The live Excel file is not in this repo yet. This page is the design an FP&A partner walks with operations — plan vs reforecast, not a headcount dump.

For a working workbook, start with the [FP&A variance dashboard](https://github.com/saisiri-bandaru/fpna-variance-dashboard) or the [13-week cash forecast](https://github.com/saisiri-bandaru/thirteen-week-cash-forecast).

## Business question

If volume is −10% next quarter, how many roles do we hold, how many do we backfill, and what does loaded cost do to OpEx?

## Drivers

| Driver | Why it matters |
| --- | --- |
| Production or transaction volume | Sets required productive hours |
| Hours per unit / occupancy | Converts volume to FTE |
| Vacancy and time-to-fill | Stops the model from assuming a hire on day one |
| Loaded cost (wage + benefits + OT) | Turns FTE into the P&L line |
| Timing of the reforecast | Plan vs current outlook bridge |

## Sample plan vs reforecast (illustrative)

| Line | Plan FTE | Reforecast FTE | Loaded $ plan | Loaded $ RF |
| --- | ---: | ---: | ---: | ---: |
| Plant ops | 120 | 112 | $9.6m | $9.0m |
| Quality | 18 | 18 | $1.6m | $1.6m |
| Warehouse | 40 | 36 | $2.6m | $2.4m |
| G&A support | 22 | 22 | $2.4m | $2.4m |
| **Total** | **200** | **188** | **$16.2m** | **$15.4m** |

Bridge story: volume-driven ops and warehouse holds, quality and G&A unchanged. Savings are timing and vacancy, not a blanket freeze.

## What I would put in the workbook

1. Assumptions tab — volume index, productivity, vacancy, load rate
2. Monthly FTE by cost center, plan and reforecast
3. Loaded-cost roll-forward into OpEx
4. One-page bridge: volume / mix / rate / timing

No employer headcount is published here.
