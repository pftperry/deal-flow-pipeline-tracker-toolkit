# Deal Flow Pipeline Tracker Toolkit

Post Fiat Analytical Infrastructure Library, v1.0. This is a reusable Excel workbook for lower-middle-market private equity deal sourcing. It tracks bankers, brokers and other intermediaries, ranks a live deal pipeline with a transparent 0-100 priority score, and flags overdue deal next actions and lapsed intermediary touches on one dashboard. Every calculated cell is a live formula.

## Files

| File | Purpose |
|---|---|
| `Deal_Flow_Pipeline_Tracker_Toolkit.xlsx` | The workbook: 7 tabs, values recalculated and cached in Microsoft Excel |
| `build_deal_flow_tracker.py` | Reproducible openpyxl build script: `python build_deal_flow_tracker.py out.xlsx` |
| `excel_recalc_and_screenshots.ps1` | Excel COM script that runs CalculateFull, scans each tab for formula errors, saves and exports the screenshots |
| `evidence/*.png` | Excel-rendered screenshots (dashboard, pipeline inputs and formula outputs, follow-up tracker, worked example, validator checklist) |

Workbook SHA-256: `D1DFBB7B03831EAF1B8469B012E32D50BE55D649E972833887FDBF2B1D587780` (52414 bytes)

## Tabs

1. ReadMe - Validator Instructions (Excel forbids `/` in tab names)
2. Settings and Scoring Rules: weights, stage probabilities, process scores, tier cadences and thresholds
3. Intermediary Directory: 25 rows of anonymized intermediaries, with live roll-ups of deals sourced, active deals and best priority score
4. Deal Pipeline Log: 40 rows covering stage, next action and due date, plus the priority score, rank and follow-up flag
5. Follow-Up Cadence Tracker: intermediary touch cadence by tier and every deal's next action, with overdue flags
6. Ranking Dashboard: KPIs, top 5 deals, stage funnel, most overdue items and a one-line dashboard output
7. Anonymized Worked Example: a live snapshot, 4 named judgment calls and the dashboard output

Colour legend: yellow cells are inputs, blue cells are formulas and green cells are key outputs. The worked example uses a fixed as-of date (`Settings!B4` = 2026-10-04) so it is reproducible. For live use, set it to `=TODAY()`.

## Scoring and flag logic

- **Priority score (0-100)** = 100 x (30% stage probability + 30% thesis fit + 15% EBITDA size fit + 15% process quality + 10% relationship strength). Each component is scaled from 0 to 1. Fit and relationship use (value - 1) / 4. Size fit is 1 inside the $3-15M band, 0.5 within 50% outside it, and 0 otherwise. Process scores are proprietary 1.0, limited auction 0.6 and broad auction 0.2. The formula is in `Deal Pipeline Log!R4:R43`, and the weights are in `Settings!B11:B15`, which must sum to 100% (`C16`).
- **Rank** is calculated for active deals only, with a deterministic tie-break on row order (`S:T`).
- **Deal follow-up flag** (`W`) is `Inactive - <stage>`, `NO NEXT ACTION`, `OVERDUE Nd`, `DUE IN Nd` (within the 7-day window), `STALE` (no activity for more than 21 days) or `On track`.
- **Intermediary cadence flag** (`Follow-Up Cadence Tracker!I5:I29`): the next touch is due on the last contact date plus the tier cadence (A = 14 days, B = 30, C = 60).
- **Pipeline health** (`Ranking Dashboard!B15`) is based on overdue deal actions plus overdue intermediary touches. Zero gives HEALTHY, 1-3 gives ATTENTION, and more than 3 gives AT RISK.

## Validator instructions: criterion -> tab -> cell range

| # | Acceptance criterion | Tab | Cell range | Expected (worked example) |
|---|---|---|---|---|
| 1 | Validator Read Me | ReadMe - Validator Instructions | `A1:E43`; checklist `A26:E36`; FORMULATEXT reference `C39:C43` | 7 tabs |
| 2 | Anonymized intermediary directory | Intermediary Directory | `A4:M28` (inputs `A:J`, roll-ups `K:M`) | 10 intermediaries, code names only |
| 3 | Deal pipeline log with stage and next action | Deal Pipeline Log | stage `G4:G43`, next action `K4:L43` | 12 deals, 10 active |
| 4 | Live priority score and rank | Deal Pipeline Log | `R4:T43`; weights `Settings!B11:C16` | Rank 1 Project Falcon 91.0 |
| 5 | Overdue follow-up flags | Deal Pipeline Log / Follow-Up Cadence Tracker | `W4:W43`; `I5:I29`; `H33:H72` | 4 overdue deal actions, 4 overdue intermediary touches |
| 6 | Follow-up cadence tracker | Follow-Up Cadence Tracker | `A5:K29` (intermediaries), `A33:H72` (deals) | Banker I 66 days overdue |
| 7 | Summary / ranking dashboard | Ranking Dashboard | KPIs `A4:C15`, top 5 `A19:F23`, funnel `A27:E37`, output `B44` | AT RISK - 8 overdue follow-ups |
| 8 | Anonymized worked example | Anonymized Worked Example | snapshot `A8:B16`, judgment calls `A20:G23`, output `A26` | 4 judgment calls |
| 9 | Visible live formulas | all | e.g. `Deal Pipeline Log!R4`, `W4`; `ReadMe!C39:C43` | 1,343 formula cells |
| 10 | Zero formula errors | all | `excel_recalc_and_screenshots.ps1` scan | 0 error cells |

## Worked example: Sponsor X (anonymized, scaled, illustrative)

This is Sponsor X's $3-15M EBITDA services and distribution pipeline: 10 intermediaries and 12 opportunities, as of 2026-10-04.

| Rank | Deal | Score | Stage | Intermediary | Flag |
|---:|---|---:|---|---|---|
| 1 | Project Falcon | 91.0 | LOI Submitted | Operator H | DUE IN 2d |
| 2 | Project Atlas | 79.0 | Management Meeting | Banker A | DUE IN 4d |
| 3 | Project Cedar | 63.5 | CIM Review | Broker C | OVERDUE 8d |
| 4 | Project Juniper | 62.0 | IOI Submitted | Attorney G | DUE IN 5d |
| 5 | Project Iris | 61.5 | Sourced | Founder J | On track |

**Dashboard output (`Ranking Dashboard!B44`):** Top priority: Project Falcon (score 91.0, LOI Submitted, via Operator H). 10 active of 12 deals; probability-weighted EBITDA $21.98M. Follow-ups: 4 overdue deal actions, 4 overdue intermediary touches, 5 due within 7 days. Most overdue: Project Granite (19d) and Banker I (66d). Health: AT RISK - 8 overdue follow-ups; clear backlog before new outreach.

Named judgment calls:
- JC1: Project Ember ($18.0M, above the band) is kept with half size credit.
- JC2: Project Birch's thesis fit is held at 3 because it is a broad auction with an ask above our view.
- JC3: Passed and lost deals are kept for intermediary attribution but excluded from the ranking.
- JC4: The as-of date is fixed instead of using `TODAY()`, so the example is reproducible.

## Verification

Recalculated in Microsoft Excel through COM (`CalculateFull`). The `UsedRange.SpecialCells(xlCellTypeFormulas, xlErrors)` scan found **0 error cells on all 7 tabs** (1,343 formula cells). Cached values were saved.

All data is anonymized and illustrative: code names only, with no firm, person or company identifiers.
