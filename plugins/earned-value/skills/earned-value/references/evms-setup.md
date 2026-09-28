# Setting up an EVMS — Australian contractor practice

Aligned to AS 4817-2006 / AS ISO 21508 and general Australian contractor practice. Every organisation sets its own procedure; treat the structures below as illustrative, not as any company's system. Read this when the user is establishing EVM on a project rather than calculating on existing data.

Contents: [When EVM applies](#when-evm-applies) · [The 11 steps](#the-11-steps) · [Cost code structure](#cost-code-structure) · [What sits in the PMB](#what-sits-in-the-pmb) · [Month-end cycle](#month-end-cycle) · [Integrity checks](#integrity-checks) · [Baseline maintenance](#baseline-maintenance) · [Reporting outputs](#reporting-outputs)

## When EVM applies

Each organisation sets its own thresholds. The usual reasons to run full EVM are:

- The job is large enough that a small percentage overrun is a big dollar number
- The duration is long enough for trends to show and for action to make a difference
- There's a significant design component, so scope moves while work is under way
- The contract requires earned value reporting
- There are many subcontract packages with dependencies between them

Short, simple or service-type contracts rarely justify it.

## The 11 steps

**Tender phase (1–3)** — done by planning and estimating, before award.

1. **Define the WBS and cost codes.** Deliverable-oriented, not resource-oriented. Product before process. Established before any delivery activity, including approved early works. Coded consistently into the estimate and the programme.
2. **Assign responsibility.** Every WBS element gets a named owner, captured in the responsibility assignment matrix and the authorisation matrix.
3. **Schedule the work.** Activity detail must match how progress and cost will be recorded. Activity durations no longer than one reporting period, or progress becomes guesswork.

**Establishment phase (4–6)** — after award, before delivery.

4. **Develop the time-phased budget.** Budget allocated to activities in line with when the cost will actually be incurred, producing the planned value curve.
5. **Assign objective measures of performance.** One technique per work package, chosen before work starts, recorded in the WPM document.
6. **Set the performance measurement baseline.** Must reconcile to the signed tender summary. Below-the-line items stay out.

**Delivery phase (7–8)**

7. **Authorise and perform the work.** Set the variance threshold here — typically ±10%, tapered tighter as the project progresses and reaction time shrinks.
8. **Record and report performance data.** The month-end cycle below.

**Review and action (9–11)**

9. **Analyse performance data.** Variance analysis report against threshold.
10. **Take management action.** Corrective actions with named owners, tracked on an action register through regular project reviews.
11. **Maintain the baseline.** Controlled change only.

## Cost code structure

Use your organisation's own coding scheme. Whatever it is, it should be hierarchical, carry
location and type of work, and stay consistent from job to job so benchmarks mean something.
The example below is illustrative only:

```
IND         Indirect costs
  IND-MOB     Mobilisation and set-up
  IND-TRC     Time-related costs (site team, facilities, plant on hire)
  IND-DEM     Demobilisation
  IND-OTH     Other indirects

DIR         Direct costs
  DIR-[area]            Location: zone, building, level or structure
  DIR-[area]-[trade]    Type of work: e.g. DIR-Z02-CONC = Zone 2 concrete
  (add a further level only if it's agreed before codes are issued)

ALW         Allowances, contingency and reserve (outside the PMB)
  ALW-DES     Design development allowance
  ALW-RSK     Risk contingency
  ALW-ESC     Escalation
  ALW-MR      Management reserve
```

Keep every code at the same depth and format so reports sort and roll up cleanly.

## What sits in the PMB

```
Contract price
├── Project budget
│   ├── Performance Measurement Baseline  ← EVM operates here only
│   │   ├── Indirect cost codes
│   │   └── Direct cost codes
│   ├── Allowances & contingencies        ← no PV, EV or AC recorded
│   └── Management reserve                ← released by sponsor only
└── Margin
```

No earned value is ever recorded against allowances or contingency. Contingency is never used to absorb a cost overrun — that hides the variance the system exists to expose. Transfers happen only on an approved baseline change / budget transfer request.

## Month-end cycle

The cut-off for schedule progress must be the same as the cut-off for cost, or EV and AC describe different weeks.

| When | Action | Typical owner |
|---|---|---|
| A few days before cut-off | Close budget transfers; confirm PV in the programme | Planner |
| Cut-off | Status the programme; enter timesheets and ledger costs | Planner / site admin |
| Straight after cut-off | Run a draft CPR **before the ledger closes** | Project controls |
| Next 1–2 days | Review cost allocations and accruals; confirm EV | Project controls / accountant |
| Then | Load final AC and EV; update the forecast; issue the CPR | Project controls |
| Shortly after | Issue the variance report with causes, actions and owners | Project manager |
| Within the month | Project review, then escalation as your governance requires | Project leadership |

The 1st draft CPR before ledger close is the point of the whole cycle — it is a data-integrity tool, not a report. It surfaces mis-allocated costs and missing accruals while there is still time to fix them.

## Integrity checks

Run these before any CPR is issued. Each one catches a specific, common failure:

- **EV recorded, no AC** → cost mis-allocated to the wrong code, or an accrual is missing.
- **AC recorded, no EV** → progress under-claimed, or costs booked to the wrong code.
- **PV in the period, no EV** → a real delay, or the schedule was not statused.
- **Any EV or AC against an allowance or contingency code** → coding error. Always wrong.
- **Cost code 100% complete with forecast cost to complete remaining** → the outstanding cost belongs in accruals.
- **Cost code 100% complete where EV ≠ budget** → measurement error.
- **FFC below AC** → impossible; forecast must be at least what has been spent.
- **Sharp reversal against last month with no event behind it** → last month's data was wrong, this month's, or both.

## Baseline maintenance

Changes come from outside (client variations, affecting contract value) or inside (approved methodology change, transfer from allowances). Budget for completed work is never changed. Re-baselining a whole project needs senior approval and should be rare — each re-baseline destroys the trend data that makes EVM useful.

## Reporting outputs

Four artefacts per month:

1. **Cost Performance Report (CPR)** — by cost code: budget, current period and cumulative PV/EV/AC, variances, indices, ETC/EAC/VAC, IEAC.
2. **Variance Analysis Report (VAR)** — every code breaching threshold, split negative and positive, each with cause, impact, corrective action and named owner.
3. **Progress S-curve** — cumulative PV, EV and AC, with the forecast curve beyond the data date.
4. **Planning and controls dashboard** — CPI/SPI trend, key dates status, budget vs time, risks and opportunities.
