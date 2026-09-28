# EVM formula reference

Contents: [The four inputs](#the-four-inputs) · [Current performance](#current-performance) · [Forecasting cost](#forecasting-cost) · [Choosing an EAC](#choosing-an-eac) · [TCPI](#tcpi) · [Forecasting time](#forecasting-time) · [Measurement techniques](#measurement-techniques) · [Nine-box interpretation](#nine-box-interpretation) · [Worked example](#worked-example)

## The four inputs

| Term | Also called | Definition |
|---|---|---|
| **BAC** — Budget at Completion | — | Total authorised budget for the work. Equals the performance measurement baseline. Excludes margin and management reserve. |
| **PV** — Planned Value | BCWS | Authorised budget for the work *scheduled* to be done by the data date. |
| **EV** — Earned Value | BCWP | Budget authorised for the work *actually done* by the data date. |
| **AC** — Actual Cost | ACWP | Costs actually incurred for that work, including accruals. |

All three must be measured at the same data date, on the same scope, and on the same cost basis (all excluding GST, or all including — never mixed). The single most common error in a real CPR is an AC that includes GST against an EV that does not.

## Current performance

```
CV  = EV − AC        Cost variance ($).      Negative = over budget.
CV% = CV / EV        Negative = over budget.
CPI = EV / AC        < 1.0 = over budget.

SV  = EV − PV        Schedule variance ($).  Negative = behind.
SV% = SV / PV        Negative = behind.
SPI = EV / PV        < 1.0 = behind.

% complete = EV / BAC
% spent    = AC / BAC
% variance = % complete − % spent
```

Guard every denominator — a cost code with no EV or no PV yet is normal early in a job. `=IF(AC=0,0,EV/AC)`.

SV is expressed in dollars, not time. It says how much budgeted work is missing, not how many weeks late the job is. Two projects with identical SV can be one week and six months late.

## Forecasting cost

```
ETC = EAC − AC                      Money still needed.
VAC = BAC − EAC                     Negative = forecast overrun.
VAC% = VAC / BAC
```

## Choosing an EAC

The formula encodes a judgement about the future. Pick it deliberately and state which one was used.

| Formula | Assumption | When it is the honest choice |
|---|---|---|
| `EAC = AC + (BAC − EV)` | Remaining work runs to budget | The variance was a one-off event (a storm, a single design change). Also standard below ~15–20% complete, where indices are too volatile to extrapolate. |
| `EAC = BAC / CPI` | Cost performance to date continues | Systemic cause — under-priced rates, a productivity shortfall, a labour market. The default independent check. |
| `EAC = AC + [(BAC − EV) / CPI]` | Same, but from the current point | Same as above, but the project has been re-baselined. Prefer this over `BAC/CPI` when the baseline has moved. |
| `EAC = AC + [(BAC − EV) / (CPI × SPI)]` | Cost *and* schedule performance both persist | Behind and overspent, with time-related overheads accruing. Produces the highest, usually most pessimistic outturn. |
| `EAC = AC + new ETC` | The original basis was wrong | The remaining scope has been genuinely re-estimated bottom-up. Nothing to calculate; document the basis. |

An **IEAC** (independent EAC) is any of these calculated from the data as a challenge to the team's own forecast final cost. Below roughly 30% complete an IEAC is unreliable — CPI has not stabilised — so suppress it rather than publish a number that will swing wildly.

The gap between FFC and IEAC is the most useful line in a cost report. It is the size of the recovery the team is implicitly promising.

## TCPI

```
TCPI (to BAC) = (BAC − EV) / (BAC − AC)
TCPI (to EAC) = (BAC − EV) / (EAC − AC)
```

The cost efficiency the remaining work must achieve to land on the target. Compare it to the CPI achieved so far — that comparison is the whole point.

- TCPI ≈ CPI: the target is consistent with performance.
- TCPI > CPI: recovery is required. A gap beyond about 10% (TCPI ≳ 1.10 against a CPI near 1.0) is generally not achievable, and the target should be reset rather than defended.
- If AC exceeds BAC the denominator goes negative and TCPI is meaningless — the budget is spent; report it that way.

## Forecasting time

Traditional EV forecasts cost, not time. SPI returns to 1.0 at completion regardless of how late the project is, because EV converges on BAC — so late in a job it flatters the picture badly.

**Earned Schedule (ES)** fixes this by measuring where on the *time* axis the current EV was planned to be achieved:

```
ES     = the time at which the current cumulative EV was planned to be earned
SV(t)  = ES − AT          (AT = actual time elapsed)
SPI(t) = ES / AT
EAC(t) = PD / SPI(t)      (PD = planned duration)
```

ES stays meaningful to the end of the project. Where the data allows it, prefer SPI(t) over SPI. Otherwise say plainly that the schedule position comes from the updated programme and critical path, not from SPI.

## Measurement techniques

| Technique | Method | Use for | Watch |
|---|---|---|---|
| Units complete | Physical count × EV rate/unit | Repetitive quantified work — concrete, pipe, panels | Needs a reliable budget quantity |
| Milestone weighting | Budget split across milestones, earned on full completion | Long packages, procurement, subcontracts | Milestones must be defined up front and be genuinely binary |
| Fixed formula (0/100, 20/80, 50/50) | Fixed split at start and finish | Activities inside one reporting period | Never for long activities — 50/50 hands over half the budget for turning up |
| % complete | Assessed percentage × budget | Where nothing else fits | Subjective; back it with weighted steps and an auditable basis |
| Level of effort | EV = PV automatically | Supervision, site management, time-related overhead | Can never show a schedule variance. Keep to a minimum. |

Rules that hold across all of them: EV for 100% complete work equals its budget exactly; progress is never accrued; and no EV is recorded against allowances or contingency.

## Nine-box interpretation

|  | SV > 0, SPI > 1 | SV = 0, SPI = 1 | SV < 0, SPI < 1 |
|---|---|---|---|
| **CV > 0, CPI > 1** | Ahead, under budget | On schedule, under budget | Behind, under budget |
| **CV = 0, CPI = 1** | Ahead, on budget | On plan | Behind, on budget |
| **CV < 0, CPI < 1** | Ahead, over budget | On schedule, over budget | Behind, over budget |

"Behind but under budget" is the box that most often hides a data problem: under-claimed progress produces exactly this signature. Check EV before celebrating the CV.

## Worked example

BAC $10.0M · at the data date PV $4.0M, EV $3.4M, AC $3.9M.

```
CV  = 3.4 − 3.9 = −$0.5M          CPI = 3.4 / 3.9 = 0.87
SV  = 3.4 − 4.0 = −$0.6M          SPI = 3.4 / 4.0 = 0.85
% complete = 34%                  % spent = 39%
```

Behind and overspent. At 34% complete the indices carry weight, so the systemic EAC applies:

```
IEAC = 10.0 / 0.87 = $11.5M       VAC = 10.0 − 11.5 = −$1.5M
TCPI (to BAC) = (10.0 − 3.4) / (10.0 − 3.9) = 1.08
```

Landing on budget now requires 1.08 efficiency against 0.87 achieved — a 24% lift on everything remaining. That is not a corrective action, it is a re-baseline. If the team's FFC is $10.4M, the reportable finding is the $1.1M gap between their forecast and the independent estimate, and the question of what specifically changes to close it.
