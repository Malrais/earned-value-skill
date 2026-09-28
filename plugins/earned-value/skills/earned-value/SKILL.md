---
name: earned-value
description: "Earned value management for construction and infrastructure projects — calculating and interpreting PV, EV, AC, CV, SV, CPI, SPI, EAC, ETC, VAC, TCPI and IEAC; building Work Progress Measurement (WPM) sheets and Cost Performance Reports (CPR) with live formulas and S-curves; writing variance analysis commentary; and setting up an EVM system aligned to AS 4817 / ISO 21508. Use this skill whenever the user mentions earned value, EVM, EVMS, CPR, WPM, cost performance report, variance analysis report, S-curve, progress claim analysis, forecast final cost, cost/schedule performance index, budget at completion, performance measurement baseline, cost codes with budgets and actuals, or asks why a project is behind or overspent — even when they don't use the words 'earned value'. Also use it when the user uploads a cost report, progress spreadsheet, or cost/budget/actuals table and wants performance assessed or forecast."
---

# Earned Value Management

EVM answers three questions from three numbers: how much work was *planned* by now (PV), how much was *achieved* (EV), and what it *cost* (AC). Everything else is arithmetic on those three plus the budget (BAC).

The common failure in practice is not the arithmetic — it is comparing the wrong things. Comparing AC to PV (spend vs budget) tells you nothing about performance, because it ignores how much work was actually done. Always anchor on EV.

## Choose the mode first

| The user wants | Do this |
|---|---|
| A number or an interpretation from figures they gave | Calculate inline. Show the working. No file. |
| A working tracker, CPR, WPM sheet, or S-curve | Run `scripts/build_evm_workbook.py` (see below) |
| Their own cost report or spreadsheet assessed | Run `scripts/analyse_evm.py` — never eyeball a real CPR |
| To know why a project looks wrong, or a figure that won't reconcile | Read `references/diagnostics.md` |
| To set up EVM on a project from scratch | Read `references/evms-setup.md` and walk the 11 steps |
| To understand a formula or which EAC to use | Read `references/formulas.md` |

## Core formulas

Read `references/formulas.md` for the full set with interpretation rules, EAC selection guidance, and Earned Schedule. The working minimum:

```
CV  = EV − AC          CPI = EV / AC          % complete = EV / BAC
SV  = EV − PV          SPI = EV / PV          % spent    = AC / BAC
CV% = CV / EV          SV% = SV / PV
VAC = BAC − EAC        ETC = EAC − AC         TCPI = (BAC − EV) / (BAC − AC)
```

Negative variance and index below 1.0 are both bad news, on either axis.

## Analysing someone else's cost report

Never read a production CPR by eye. Run it through the analyser, which handles the traps
that make a manual read confidently wrong:

```bash
python scripts/analyse_evm.py --data cpr.xlsx --sheet "CPR" --header-row 6
```

It matches column headings loosely, so the usual names work unchanged. Where a sheet repeats
PV/EV/AC for current period, cumulative and cumulative-excluding-GST — which production
reports do — it refuses to guess, lists the candidate columns, and asks for explicit ones:

```bash
python scripts/analyse_evm.py --data cpr.xlsx --sheet CPR --header-row 6 \
  --col bac=E --col pv=S --col ev=T --col ac=U --col ffc=AD
```

Cumulative is what a monthly report is built on. The script drops rows holding Excel error
values rather than reading them as zero, skips total rows so nothing is double-counted,
measures planned-value coverage on non-zero values, and returns integrity flags, the full EAC
family, the worst variances by dollars, and written findings. Read `references/diagnostics.md`
before interpreting anything unusual.

## Interpreting, not just reporting

A CPI of 0.94 is not a finding. What it means for the final cost is. Whenever you report indices, carry them through to money and time:

- **Forecast the outturn.** State the independent EAC (BAC/CPI) alongside the team's own forecast, and name the gap. A forecast final cost that sits below the independent EAC is a claim that performance will improve — say so, and ask what will change.
- **Say whether recovery is credible.** TCPI is the test. Above about 1.10 the required efficiency is generally not achievable, and the honest conclusion is that the budget is gone, not that the team must try harder.
- **Gate every index on maturity.** Below 15% complete the indices are noise and must not be forecast from; below 30% no independent EAC is publishable; past 70% SPI drifts to 1.0 regardless of lateness; past 90% TCPI swings on a tiny remainder. `references/diagnostics.md` has the full table. Use a tolerance band matched to the project's variance threshold rather than a bare `CPI < 1` test — calling 0.9998 an overrun destroys credibility.
- **Watch the level-of-effort share.** LoE earns value equal to planned value by definition, so it can never show a variance and it dilutes every index it sits in. A CPI of exactly 1.00 across a large share of budget is a signature, not a result. Report direct-work CPI separately when LoE exceeds roughly a quarter of the PMB.
- **Distrust suspiciously clean lines.** EV recorded with no matching AC usually means costs are unallocated or unaccrued. AC with no EV usually means progress is under-claimed. Both are data problems that will reverse next month, and both should be raised before any conclusion is drawn from the variance.
- **Separate buying gain from performance.** A positive CV on a lump-sum subcontract package is often a procurement result banked at award, not efficiency during delivery. It will not repeat on the next package.

## Building the workbook

`scripts/build_evm_workbook.py` produces a complete, formula-live workbook: a WPM sheet where progress is entered, a CPR that rolls up by cost code, a time-phased budget feeding an S-curve, a dashboard with CPI/SPI trend, a variance analysis report driven by a threshold, and a budget transfer log.

```bash
python scripts/build_evm_workbook.py \
  --project "Project Name" \
  --out /mnt/user-data/outputs/project_evm.xlsx \
  --data cost_codes.csv \
  --periods 18 --threshold 0.10
```

`--data` takes a CSV with the columns `cost_code, description, wbs_l1, wbs_l2, responsibility, budget_qty, unit, budget_value, technique` (any missing column is left blank for the user to fill). Without `--data` the script writes a small worked example so the user can see the expected format. Run `python scripts/build_evm_workbook.py --help` for the rest.

The workbook carries defined names (`PMB_BAC`, `PMB_EV`, `PMB_AC`, `CPR_DATA`, `WPM_DATA`,
`TP_DATA` and others) so anything reading it later can find the data block and the totals
without guessing where the rows stop. The CPR runs the integrity checks per cost code in a
Data Check column, and the Status column reports `CHECK DATA` ahead of any variance flag,
because a threshold breach on bad data wastes a month.

After generating, always recalculate before delivering — openpyxl writes formulas with no cached values, so an unrecalculated file looks empty in most previewers:

```bash
python /mnt/skills/public/xlsx/scripts/recalc.py /mnt/user-data/outputs/project_evm.xlsx
```

Then `present_files` the workbook. Tell the user which cells they fill: blue-font cells are inputs, everything black is calculated.

## Measuring earned value honestly

The technique must be picked *before* work starts, or EV becomes an opinion. `references/formulas.md` covers each; the short version:

- **Units complete** — physical count against a known budget quantity. Best when available.
- **Milestone weighting** — budget assigned to defined milestones, earned only on full completion. For long packages.
- **Fixed formula** (0/100, 20/80, 50/50) — short activities only, inside one reporting period.
- **% complete** — subjective, so back it with weighted steps rather than a manager's judgement.
- **Level of effort** — time-related work with no output measure. EV always equals PV, so LoE can never show a schedule variance. Keep it to a practical minimum; a project heavy in LoE has an SPI that means nothing.

Never accrue progress. EV for completed work must equal its budget exactly, and no EV is recorded against allowances or contingency — those sit outside the performance measurement baseline.

## Australian contractor context

Local practice (AS 4817-2006, as commonly applied by Australian contractors) differs from the PMBOK exam framing in ways that matter:

- The **PMB is indirect + direct cost codes only**. Allowances, contingency, escalation and margin sit outside it.
- **Forecast Final Cost (FFC)** from the cost system is the working EAC; the calculated IEAC is the challenge to it, not a replacement.
- Reporting runs a fixed month-end cycle: budget transfers close, schedule statused, 1st draft CPR before ledger close, accruals reviewed, final CPR, then variance analysis report and monthly project review.
- A **±10% variance threshold** typically triggers mandatory commentary and a corrective action with a named owner.

`references/evms-setup.md` has the 11-step process, an illustrative cost code structure, and a typical month-end cycle.

## Writing the variance analysis

Commentary that says "CPI is 0.94 due to productivity issues" is worthless. Each breach needs: what the variance is in dollars, what caused it, whether it is one-off or systemic, its effect on the forecast, the corrective action, and who owns it. One-off versus systemic is the load-bearing judgement — it decides which EAC formula is honest.

Keep the language plain and direct. Australian construction reporting is blunt; write it that way.
