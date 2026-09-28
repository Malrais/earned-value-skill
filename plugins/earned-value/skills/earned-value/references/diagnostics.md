# Diagnostic patterns and worked scenarios

What the numbers mean when they misbehave, and how a project controls professional reads
each one. Every pattern below came out of testing against real cost reports and six
project archetypes, not from a textbook.

Contents: [Read the data before the variance](#read-the-data-before-the-variance) ·
[Maturity gates](#maturity-gates) · [Column traps in a real CPR](#column-traps-in-a-real-cpr) ·
[Worked scenarios](#worked-scenarios) · [Archetype notes](#archetype-notes) ·
[Questions that expose a soft forecast](#questions-that-expose-a-soft-forecast)

## Read the data before the variance

A threshold breach on bad data sends a team chasing a cause that does not exist. Clear
these first, every month, before any commentary is written.

| Pattern | Almost always means | Move |
|---|---|---|
| EV recorded, AC zero | Cost mis-coded, or an accrual missing | Chase the accrual before reporting a favourable CV |
| AC recorded, EV zero | Progress under-claimed, or cost on the wrong code | Status the activity; do not report the overrun yet |
| EV exceeds BAC | Measurement error, or budget removed without EV adjusted | Cap EV at budget and find the transfer |
| FFC below AC | Impossible — forecast cannot sit below money spent | The forecast was never updated; it is still defaulting to BAC |
| 100% complete, ETC remaining | Outstanding cost belongs in accruals | Move it, or the code will never close |
| PV in the period, no EV and no AC | Genuine delay, or the schedule was not statused | Check the programme before calling it a delay |
| EV zero across every code while AC is large | Wrong column, not a project with no progress | Re-check which PV/EV/AC block was read |
| Sharp reversal against last month with no event behind it | Last month was wrong, this month is, or both | Reconcile before publishing |

A favourable variance deserves the same scrutiny as an adverse one. On lump-sum
subcontract packages a positive CV is usually buying gain banked at award, or cost not
yet posted — neither repeats, and treating it as efficiency builds a forecast on sand.

## Maturity gates

Indices do not mean the same thing at every stage. Applying them uniformly is the most
common way a cost report loses its audience.

| Stage | CPI / SPI | IEAC | TCPI |
|---|---|---|---|
| Below 15% complete | Report, do not forecast from them — small denominators swing wildly | Suppress | Meaningless |
| 15–30% | CPI usable with caution; SPI still rough | Suppress — CPI has not settled | Directional only |
| 30–70% | Both meaningful; this is the window where EVM earns its keep | Publish | Publish |
| 70–90% | CPI still good; **SPI drifts to 1.0 regardless of lateness** — read the programme instead | Publish | Publish |
| Above 90% | CPI is effectively final | Converges on actual | Unstable — tiny denominator, ignore the number |

Use a tolerance band rather than a bare `CPI < 1` test. A CPI of 0.9998 is on budget by any
sane reading, and calling it an overrun destroys credibility. Match the band to the
project's own variance threshold so the dashboard and the CPR never disagree.

## Column traps in a real CPR

A production cost report repeats PV, EV and AC three times — current period, cumulative,
and cumulative excluding GST. Reading the wrong block produces a complete, plausible,
entirely wrong analysis.

- **Cumulative is what a monthly report is built on.** Current-period columns are often
  derived as `cumulative − previous`, so when the previous-month reference breaks they fill
  with `#REF!`.
- **An Excel error read as zero is worse than a crash.** `#REF!`, `#N/A` and `#DIV/0!`
  silently become 0 in most parsers, and the result is a report claiming no progress on a
  job that is half built. Reject error values and say how many rows were dropped.
- **A zero planned value is indistinguishable from one never entered.** Measure PV coverage
  by non-zero values, not by "the cell has something in it". If PV covers well under the
  whole budget, SPI and SV describe that subset only and should be labelled as such — or
  suppressed. On one real CPR this was the difference between SPI 2.69 and an honest
  "schedule position must come from the programme".
- **Never sum a row whose label contains "total" or "subtotal".** Double-counting a total
  row is the quiet way to overstate a project by a factor of two.
- **Check the GST basis matches.** An AC including GST against an EV excluding it produces a
  CPI near 0.91 on a project performing perfectly. This is the single most common error in a
  real cost report.

## Worked scenarios

Drawn from the test suite. Each shows the reading, not just the arithmetic.

### Early stage — metro station, 1.7% complete
`CPI 0.96 · SPI 0.29 · IEAC suppressed`

Nothing here is reportable as performance. At 1.7% complete a handful of codes drive every
index, and SPI 0.29 reflects a mobilisation sequence, not a late project. The honest report
says the baseline is set, actuals are flowing, and indices will be meaningful from around
15%. Publishing "29% schedule performance" at this stage teaches everyone to ignore the CPR.

### Midstream systemic overrun — data centre, 48.5% complete
`CPI 0.77 · SPI 0.94 · IEAC $590M against BAC $455M · TCPI 1.39`

This is the real thing. At nearly half complete the indices have settled, so `BAC / CPI`
applies: a forecast overrun of roughly $135M. TCPI 1.39 against CPI 0.77 means everything
remaining must run 80% more efficiently than everything so far. That is not a corrective
action, it is a re-baseline, and the report should say so in those words.

### Budget exhausted with work remaining — metro station, 66% complete
`CPI 0.50 · AC has passed BAC · TCPI undefined`

TCPI's denominator `BAC − AC` has gone negative, which is the system telling you the budget
is gone. There is no recovery scenario to model. The only useful output is a bottom-up
re-estimate of the remaining scope and an escalation, not an index.

### Near complete — hospital, 95% complete
`CPI 0.95 · SPI 0.96 · TCPI 42.1`

TCPI is arithmetically correct and operationally meaningless — a tiny remaining budget in
the denominator. Report CPI as effectively final, forecast the outturn from it, and drop
TCPI entirely. SPI 0.96 at 95% complete says nothing about whether the job finishes on
time; only the programme does.

### Favourable variance — commercial fitout, 51% complete
`CPI 1.26 · TCPI 0.82`

Too good to accept without testing. On a fitout this size a 26% favourable variance is
either buying gain from a well-priced package, or trade costs not yet posted. Check
subcontractor claims and accruals before letting anyone book the saving.

### Level-of-effort heavy — metro station, 47% complete
`CPI 1.00 · SPI 0.94`

A CPI of exactly 1.00 across a large share of the budget is a signature, not a result. LoE
codes earn value equal to planned value by definition, so they can never show a variance and
they dilute every index they sit in. Report the direct-work CPI separately from the LoE
share, and note what proportion of the PMB is LoE. Above roughly a quarter, the headline
indices are describing the measurement method more than the project.

### Not started — commercial fitout, 0% complete
`CPI n/a · SPI 0 · TCPI 1.00`

Every ratio divides by zero or returns a default. Say "not started" rather than publishing
zeros that look like catastrophic performance.

## Archetype notes

**Data centre.** Equipment-heavy — chillers, UPS, generation and switchgear can exceed half
the PMB, and nearly all of it is milestone-weighted on supply. Expect large EV steps on
delivery and factory acceptance, which makes monthly CPI jumpy without anything being wrong.
Commissioning and IST at the tail carry disproportionate schedule risk that cost indices
never show.

**Hospital.** Long fitout tail with many small, interdependent trades. Services (mechanical,
electrical, hydraulic, medical gas) dominate the back half. FF&E and medical equipment are
usually client-supplied in part — make sure the budget and the EV cover the same scope, or
CPI will look favourable for no reason.

**Metro / infrastructure.** Very large indirect and time-related component. LoE share is high
by nature, so the LoE dilution note applies. Excavation and diaphragm walling are excellent
units-complete candidates; tunnel interfaces and systems are milestone-weighted and lumpy.
Fixed time delay allowances drawn down monthly behave like LoE and should be watched.

**Marine / waterfront.** Weather and tide dependency makes period-on-period variance noisy.
Plant standby is a large fixed cost that continues whether or not work proceeds, so an
adverse CV often reflects lost days rather than productivity. Check the weather log before
writing the cause.

**Commercial fitout.** Short duration means fixed-formula techniques are appropriate for
most activities, and the whole job may run inside a handful of reporting periods. Indices
mature quickly but the window to act is tiny — tighten the variance threshold rather than
loosening it.

**MMC / modular.** Split factory and site. Module manufacture is genuinely units-complete and
measures beautifully; the risk sits in the stage gates and in inter-module stitching on site,
which is where rework hides. Watch for EV earned in the factory against modules that have not
passed witness testing — that is accrued progress by another name.

## Questions that expose a soft forecast

Ask these in a project review. They are more useful than any index.

1. The FFC sits below the independent EAC. What specifically changes to close that gap, who
   owns it, and by when?
2. Which variances are one-off and which are systemic? That single call decides which EAC
   formula is honest.
3. What is the TCPI, and has this team ever achieved that efficiency on this job?
4. How much of the PMB is level of effort? What is CPI on direct work alone?
5. Where accruals were estimated rather than invoiced, what happens to CPI if they are 20%
   light?
6. Which codes are 100% complete with cost still forecast, and why has that not moved to
   accruals?
7. What does the programme say about the finish date, independent of SPI?
