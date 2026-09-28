#!/usr/bin/env python3
"""Analyse an existing cost dataset: indices, forecasts, integrity flags, findings.

Use this when handed someone else's cost report rather than building a new one.
It takes a CSV or an .xlsx sheet holding one row per cost code and, at minimum,
budget, earned value and actual cost.

    python analyse_evm.py --data cpr.csv
    python analyse_evm.py --data book.xlsx --sheet "CPR-SMW" --header-row 6
    python analyse_evm.py --data cpr.csv --pct-complete-override 0.42 --json out.json

Column names are matched loosely, so the usual headings work without renaming:
    code      cost code, cost account, wbs, code
    desc      description, detail, scope, name
    bac       budget at completion, bac, budget, current budget, revised budget
    pv        planned value, pv, bcws
    ev        earned value, ev, bcwp
    ac        actual cost, ac, acwp, actual, cost to date
    ffc       forecast final cost, ffc, eac, forecast
    owner     responsibility, owner, manager

Anything it cannot match it reports rather than guessing.
"""

import argparse
import json
import re
import sys

ALIASES = {
    "code": ["cost code", "cost account", "costcode", "account", "code", "wbs", "item"],
    "desc": ["description", "detail description", "detail", "scope", "name", "activity"],
    "bac": ["budget at completion", "bac", "current budget", "revised budget",
            "budget value", "budget", "total budget", "approved budget"],
    "pv":  ["planned value", "pv", "bcws", "budgeted cost of work scheduled", "planned"],
    "ev":  ["earned value", "ev", "bcwp", "budgeted cost of work performed", "earned"],
    "ac":  ["actual cost", "ac", "acwp", "actual cost of work performed", "actual",
            "cost to date", "actuals", "spend"],
    "ffc": ["forecast final cost", "ffc", "eac", "estimate at completion", "forecast",
            "forecast cost", "final forecast"],
    "owner": ["responsibility", "owner", "manager", "accountable", "cost code owner"],
}

MATURITY_FLOOR = 0.15   # below this, indices are noise
IEAC_FLOOR = 0.30       # below this, an independent EAC is not publishable


def norm(text):
    return re.sub(r"[^a-z0-9 ]+", " ", str(text or "").lower()).strip()


def match_columns(headers):
    """Map our field names onto the sheet's actual headings.

    Returns (mapping, ambiguous) where ambiguous names fields that matched more
    than one column - common in a real CPR, which repeats PV/EV/AC for current
    period, cumulative, and cumulative excluding GST. Guessing between them
    produces a plausible report built on the wrong block, so the caller is made
    to choose instead.
    """
    found, used, ambiguous = {}, set(), {}
    cleaned = [(i, norm(h)) for i, h in enumerate(headers)]
    for field, names in ALIASES.items():
        best = None
        for alias in names:                      # exact first, longest alias wins
            for i, h in cleaned:
                if i in used or not h:
                    continue
                if h == alias:
                    best = i
                    break
            if best is not None:
                break
        if best is None:
            # Short aliases (ac, ev, pv, bac) only ever match exactly. As substrings
            # they hit BAC, EAC, VAC and IEAC, which is worse than no match.
            for alias in (a for a in names if len(a) > 3):
                for i, h in cleaned:
                    if i in used or not h:
                        continue
                    if alias in h or h in alias:
                        best = i
                        break
                if best is not None:
                    break
        if best is not None:
            found[field] = best
            used.add(best)
            hits = [i for i, h in cleaned
                    if h and any(h == a or (len(a) > 3 and (a in h or h in a))
                                 for a in names)]
            if len(hits) > 1:
                ambiguous[field] = hits
    return found, ambiguous


ERRORS = ("#REF!", "#N/A", "#VALUE!", "#DIV/0!", "#NAME?", "#NULL!", "#NUM!")


def _letter(i):
    s = ""
    i += 1
    while i:
        i, r = divmod(i - 1, 26)
        s = chr(65 + r) + s
    return s


def to_num(v):
    """Return a float, None for blanks, or the sentinel "ERR" for an Excel error.

    An error value silently read as zero is how a broken sheet produces a
    confident, wrong analysis - so it is carried through, not swallowed.
    """
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip()
    if s.upper() in ERRORS:
        return "ERR"
    s = s.replace("$", "").replace(",", "").replace("%", "")
    if s in ("", "-", "n/a", "na", "NA"):
        return None
    neg = s.startswith("(") and s.endswith(")")
    if neg:
        s = s[1:-1]
    try:
        n = float(s)
    except ValueError:
        return None
    return -n if neg else n


def load(path, sheet, header_row):
    if path.lower().endswith((".xlsx", ".xlsm")):
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb[sheet] if sheet else wb.worksheets[0]
        rows = [[c for c in r] for r in ws.iter_rows(values_only=True)]
    else:
        import csv
        with open(path, newline="", encoding="utf-8-sig") as fh:
            rows = list(csv.reader(fh))
    if not rows:
        raise SystemExit("No rows found.")

    hr = (header_row - 1) if header_row else None
    if hr is None:                                # find the widest early row
        hr = max(range(min(15, len(rows))),
                 key=lambda i: sum(1 for c in rows[i] if str(c or "").strip()))
    return rows[hr], rows[hr + 1:]


def col_index(spec):
    """Accept a spreadsheet letter (T) or a 1-based number (20)."""
    spec = str(spec).strip()
    if spec.isdigit():
        return int(spec) - 1
    n = 0
    for ch in spec.upper():
        if not ("A" <= ch <= "Z"):
            raise SystemExit(f"Bad column reference: {spec}")
        n = n * 26 + (ord(ch) - 64)
    return n - 1


def analyse(path, sheet, header_row, pct_override, overrides=None):
    headers, body = load(path, sheet, header_row)
    cols, ambiguous = match_columns(headers)

    for field, spec in (overrides or {}).items():
        if field not in ALIASES:
            raise SystemExit(f"Unknown column field '{field}'. "
                             f"Use one of: {', '.join(ALIASES)}")
        cols[field] = col_index(spec)
        ambiguous.pop(field, None)

    missing = [f for f in ("bac", "ev", "ac") if f not in cols]
    if missing:
        raise SystemExit(
            f"Could not find column(s) for: {', '.join(missing)}.\n"
            f"Headings seen: {[str(h) for h in headers if str(h or '').strip()][:25]}")

    unresolved = {f: v for f, v in ambiguous.items() if f in ("pv", "ev", "ac", "bac")}
    if unresolved:
        lines = ["Ambiguous columns - this sheet repeats these fields, so pick one:"]
        for field, hits in unresolved.items():
            opts = ", ".join(f"{_letter(i)} ({str(headers[i]).strip()[:24] or 'blank'})"
                             for i in hits)
            lines.append(f"  {field.upper():<4} matches {opts}")
        lines.append("")
        lines.append("Re-run with explicit columns, for example:")
        lines.append("  --col ev=T --col ac=U --col pv=S")
        lines.append("A CPR normally repeats PV/EV/AC for current period and cumulative;")
        lines.append("cumulative is what a monthly report is built on.")
        raise SystemExit("\n".join(lines))

    def cell(row, field):
        i = cols.get(field)
        return row[i] if (i is not None and i < len(row)) else None

    codes, skipped, errored = [], 0, 0
    for row in body:
        if not any(str(c or "").strip() for c in row):
            continue
        bac, ev, ac = (to_num(cell(row, f)) for f in ("bac", "ev", "ac"))
        if "ERR" in (bac, ev, ac):
            errored += 1
            continue
        if bac is None and ev is None and ac is None:
            skipped += 1
            continue
        label = str(cell(row, "code") or cell(row, "desc") or "").strip()
        if re.search(r"\b(total|subtotal|sum)\b", label.lower()):
            continue                              # never double-count a total row
        codes.append({
            "code": label,
            "desc": str(cell(row, "desc") or "").strip(),
            "owner": str(cell(row, "owner") or "").strip(),
            "bac": bac or 0.0,
            "pv": (lambda x: None if x == "ERR" else x)(to_num(cell(row, "pv"))),
            "ev": ev or 0.0,
            "ac": ac or 0.0,
            "ffc": (lambda x: None if x == "ERR" else x)(to_num(cell(row, "ffc"))),
        })
    if not codes:
        raise SystemExit(
            f"No usable data rows. {errored} row(s) held Excel error values "
            f"(#REF!, #N/A and the like) and {skipped} were blank. "
            f"Fix the source sheet, or point at the right columns with --col.")

    BAC = sum(c["bac"] for c in codes)
    EV = sum(c["ev"] for c in codes)
    AC = sum(c["ac"] for c in codes)
    # Schedule variance is only meaningful over the codes that actually carry a
    # planned value. Dividing total EV by a partial PV inflates SPI badly - on a
    # real CPR with PV on a third of the codes it reads near 3.0 and means nothing.
    # A zero planned value is indistinguishable from one never entered, and either
    # way SPI measured over it is meaningless - so only non-zero PV counts as cover.
    pv_codes = [c for c in codes if c["pv"]]
    PV = sum(c["pv"] for c in pv_codes)
    EV_on_pv = sum(c["ev"] for c in pv_codes)
    has_pv = bool(pv_codes)
    pv_coverage = (sum(c["bac"] for c in pv_codes) / BAC) if BAC else 0.0
    FFC = sum((c["ffc"] if c["ffc"] is not None else c["bac"]) for c in codes)
    has_ffc = any(c["ffc"] is not None for c in codes)

    pct = pct_override if pct_override is not None else (EV / BAC if BAC else 0.0)
    cpi = EV / AC if AC else None
    spi = EV_on_pv / PV if (has_pv and PV) else None

    out = {
        "source": path, "cost_codes": len(codes), "rows_skipped": skipped,
        "rows_with_errors": errored,
        "columns_matched": {
            k: (str(headers[v]).strip() if v < len(headers)
                and str(headers[v] or "").strip() else _letter(v))
            for k, v in cols.items()},
        "BAC": BAC, "PV": PV if has_pv else None, "EV": EV, "AC": AC,
        "pct_complete": pct, "pct_spent": (AC / BAC if BAC else None),
        "CV": EV - AC, "CV_pct": ((EV - AC) / EV) if EV else None,
        "SV": (EV_on_pv - PV) if has_pv else None,
        "SV_pct": ((EV_on_pv - PV) / PV) if (has_pv and PV) else None,
        "pv_coverage": pv_coverage if has_pv else 0.0,
        "pv_codes": len(pv_codes),
        "CPI": cpi, "SPI": spi,
        "FFC": FFC if has_ffc else None,
    }

    # Forecasts. Which EAC is honest depends on whether the variance is systemic;
    # report the family and say so rather than publishing one number as the answer.
    if cpi:
        out["EAC_one_off"] = AC + (BAC - EV)
        out["EAC_systemic"] = BAC / cpi
        out["EAC_from_here"] = AC + ((BAC - EV) / cpi)
        if spi:
            out["EAC_cost_and_schedule"] = AC + ((BAC - EV) / (cpi * spi))
        out["IEAC"] = (BAC / cpi) if pct >= IEAC_FLOOR else None
        out["VAC_vs_systemic"] = BAC - (BAC / cpi)
    denom = BAC - AC
    out["TCPI_to_BAC"] = ((BAC - EV) / denom) if denom > 0 else None
    out["ETC"] = (FFC - AC) if has_ffc else None

    # Integrity first - a variance on bad data wastes everyone's month.
    integrity = []
    for c in codes:
        tag = None
        if c["ev"] > 0 and c["ac"] == 0:
            tag = "EV recorded, no actual cost - costs unposted or miscoded"
        elif c["ac"] > 0 and c["ev"] == 0:
            tag = "Actual cost, no EV - progress under-claimed or wrong code"
        elif c["bac"] > 0 and c["ev"] > c["bac"] * 1.0001:
            tag = "EV exceeds budget - measurement error"
        elif c["ffc"] is not None and c["ffc"] < c["ac"]:
            tag = "Forecast below cost already spent - impossible"
        elif c["bac"] > 0 and abs(c["ev"] - c["bac"]) < 0.005 * c["bac"] \
                and c["ffc"] is not None and c["ffc"] > c["ac"] * 1.0001:
            tag = "Complete but cost still to come - belongs in accruals"
        elif c["pv"] is not None and c["pv"] > 0 and c["ev"] == 0 and c["ac"] == 0:
            tag = "Planned value with no progress and no cost - delayed or unstatused"
        if tag:
            integrity.append({"code": c["code"], "desc": c["desc"], "issue": tag,
                              "bac": c["bac"], "ev": c["ev"], "ac": c["ac"]})
    out["integrity_flags"] = integrity

    # Worst cost variances by dollars, which is what a project review argues about.
    ranked = sorted(codes, key=lambda c: (c["ev"] - c["ac"]))
    out["worst_cost_variances"] = [{
        "code": c["code"], "desc": c["desc"], "owner": c["owner"],
        "bac": c["bac"], "ev": c["ev"], "ac": c["ac"],
        "CV": c["ev"] - c["ac"],
        "CPI": (c["ev"] / c["ac"]) if c["ac"] else None,
    } for c in ranked[:10] if (c["ev"] - c["ac"]) < 0]
    out["largest_favourable"] = [{
        "code": c["code"], "desc": c["desc"], "CV": c["ev"] - c["ac"],
        "CPI": (c["ev"] / c["ac"]) if c["ac"] else None,
    } for c in reversed(ranked[-5:]) if (c["ev"] - c["ac"]) > 0]

    out["findings"] = findings(out, pct, cpi, spi, integrity, has_pv, has_ffc)
    out["columns_ambiguous"] = {k: [_letter(i) for i in v]
                                for k, v in ambiguous.items()}
    return out


def findings(o, pct, cpi, spi, integrity, has_pv, has_ffc):
    f = []
    if o.get("rows_with_errors"):
        f.append(f"{o['rows_with_errors']} row(s) were dropped because they held Excel "
                 f"error values. The source workbook has broken references - the totals "
                 f"below exclude them and are understated.")
    if o["EV"] == 0 and o["AC"] > 0:
        f.append("Earned value is zero across every code while actual cost is not. "
                 "That is almost always the wrong column, not a project with no progress "
                 "- check the EV column with --col ev=<letter>.")
    if integrity:
        f.append(f"{len(integrity)} cost code(s) carry a data integrity flag. "
                 f"Resolve these before drawing any conclusion from the variances.")
    if not has_pv:
        f.append("No planned value column found, so there is no schedule variance. "
                 "Schedule position has to come from the programme.")
    elif o.get("pv_coverage", 1.0) < 0.9:
        f.append(f"Planned value is non-zero on only {o['pv_codes']} code(s), covering "
                 f"{o['pv_coverage']:.0%} of budget. SPI and SV describe that subset alone "
                 f"- not a project-wide schedule position. Treat SPI "
                 f"{('%.2f' % o['SPI']) if o.get('SPI') else 'n/a'} as unreliable and read "
                 f"the programme instead.")
    if cpi is None:
        f.append("No actual cost recorded - cost performance cannot be assessed.")
    elif pct < MATURITY_FLOOR:
        f.append(f"At {pct:.1%} complete the indices are still noise "
                 f"(CPI {cpi:.2f}). Report them, but do not forecast from them.")
    else:
        band = "over budget" if cpi < 0.98 else ("favourable" if cpi > 1.02 else "on budget")
        f.append(f"CPI {cpi:.2f} - {band} at {pct:.1%} complete.")
        if cpi > 1.02:
            f.append("A favourable CPI on lump-sum packages is usually buying gain banked "
                     "at award, or cost not yet posted. Confirm which before relying on it.")
    pv_ok = o.get("pv_coverage", 1.0) >= 0.9
    if spi is not None and pv_ok:
        if pct > 0.7:
            f.append(f"SPI {spi:.2f}, but past 70% complete SPI drifts to 1.0 regardless "
                     f"of lateness. Read the programme, not this number.")
        elif pct >= MATURITY_FLOOR:
            f.append(f"SPI {spi:.2f} - {'behind' if spi < 0.98 else 'on or ahead of'} plan "
                     f"in dollar terms. SV is dollars of missing work, not weeks.")
    if o.get("IEAC"):
        f.append(f"Independent EAC ${o['IEAC']:,.0f} against BAC ${o['BAC']:,.0f} "
                 f"- forecast overrun ${-o['VAC_vs_systemic']:,.0f} if performance holds."
                 if o["VAC_vs_systemic"] < 0 else
                 f"Independent EAC ${o['IEAC']:,.0f}, at or under BAC.")
        if has_ffc and o.get("FFC"):
            gap = o["IEAC"] - o["FFC"]
            if gap > 0.02 * o["BAC"]:
                f.append(f"The team's forecast sits ${gap:,.0f} below the independent "
                         f"estimate. That gap is the recovery being promised - ask what "
                         f"specifically delivers it.")
    elif cpi:
        f.append(f"Below {IEAC_FLOOR:.0%} complete, so no independent EAC is published "
                 f"- CPI has not settled enough to extrapolate.")
    t = o.get("TCPI_to_BAC")
    if t is None:
        f.append("Actual cost has reached or passed the budget with work remaining. "
                 "TCPI is undefined - the budget is gone, and the target needs resetting.")
    elif cpi and pct >= MATURITY_FLOOR:
        if pct > 0.9:
            f.append(f"TCPI {t:.2f}, but past 90% complete it swings on a small remainder.")
        elif t > cpi * 1.10:
            f.append(f"TCPI {t:.2f} against CPI {cpi:.2f} - landing on budget needs a "
                     f"{(t / cpi - 1):.0%} lift on everything left. That is a re-baseline, "
                     f"not a corrective action.")
        elif t > 1.02:
            f.append(f"TCPI {t:.2f} - recovery required, and it is within reach if the "
                     f"causes are addressed.")
    return f


def render(o):
    L = []
    m = lambda v: "n/a" if v is None else f"${v:,.0f}"
    i = lambda v: "n/a" if v is None else f"{v:.2f}"
    p = lambda v: "n/a" if v is None else f"{v:.1%}"

    L.append("=" * 68)
    L.append("EARNED VALUE ANALYSIS")
    L.append("=" * 68)
    L.append(f"Source: {o['source']}   cost codes: {o['cost_codes']}")
    if o.get("rows_with_errors"):
        L.append(f"WARNING: {o['rows_with_errors']} row(s) dropped - Excel error values")
    L.append("Columns read: " + ", ".join(f"{k}={v}" for k, v in
                                          o["columns_matched"].items()))
    if o.get("pv_coverage") is not None and 0 < o["pv_coverage"] < 0.9:
        L.append(f"NOTE: planned value covers {o['pv_coverage']:.0%} of budget "
                 f"({o['pv_codes']} codes) - SPI/SV describe that subset only")
    L.append("")
    L.append(f"  BAC {m(o['BAC']):>16}     PV {m(o['PV']):>16}")
    L.append(f"  EV  {m(o['EV']):>16}     AC {m(o['AC']):>16}")
    L.append(f"  % complete {p(o['pct_complete']):>9}     % spent {p(o['pct_spent']):>11}")
    L.append("")
    L.append(f"  CV  {m(o['CV']):>16} ({p(o['CV_pct'])})"
             f"     CPI {i(o['CPI']):>6}")
    L.append(f"  SV  {m(o['SV']):>16} ({p(o['SV_pct'])})"
             f"     SPI {i(o['SPI']):>6}")
    L.append("")
    L.append("  Forecast family")
    for k, label in (("EAC_one_off", "one-off variance   (AC + BAC - EV)"),
                     ("EAC_systemic", "systemic           (BAC / CPI)"),
                     ("EAC_from_here", "systemic from here (AC + (BAC-EV)/CPI)"),
                     ("EAC_cost_and_schedule", "cost + schedule    (CPI x SPI)")):
        if o.get(k) is not None:
            L.append(f"    {label:<40} {m(o[k]):>16}")
    L.append(f"    {'team forecast (FFC)':<40} {m(o.get('FFC')):>16}")
    L.append(f"    {'TCPI to BAC':<40} {i(o.get('TCPI_to_BAC')):>16}")

    if o["integrity_flags"]:
        L.append("")
        L.append(f"  DATA INTEGRITY - {len(o['integrity_flags'])} flag(s)")
        for fl in o["integrity_flags"][:15]:
            L.append(f"    {fl['code']:<12} {fl['issue']}")
            L.append(f"    {'':<12} BAC {m(fl['bac'])}  EV {m(fl['ev'])}  AC {m(fl['ac'])}")
        if len(o["integrity_flags"]) > 15:
            L.append(f"    ... and {len(o['integrity_flags']) - 15} more")

    if o["worst_cost_variances"]:
        L.append("")
        L.append("  WORST COST VARIANCES")
        for c in o["worst_cost_variances"]:
            L.append(f"    {c['code']:<12} {c['desc'][:34]:<34} "
                     f"CV {m(c['CV']):>14}  CPI {i(c['CPI'])}")

    if o["largest_favourable"]:
        L.append("")
        L.append("  LARGEST FAVOURABLE (check for buying gain or unposted cost)")
        for c in o["largest_favourable"]:
            L.append(f"    {c['code']:<12} {c['desc'][:34]:<34} "
                     f"CV {m(c['CV']):>14}  CPI {i(c['CPI'])}")

    L.append("")
    L.append("  FINDINGS")
    for line in o["findings"]:
        L.append(f"    - {line}")
    L.append("=" * 68)
    return "\n".join(L)


def main():
    ap = argparse.ArgumentParser(description="Analyse an existing EVM dataset.")
    ap.add_argument("--data", required=True, help="CSV or .xlsx to analyse")
    ap.add_argument("--sheet", help="Worksheet name (xlsx only)")
    ap.add_argument("--header-row", type=int,
                    help="1-based header row; auto-detected when omitted")
    ap.add_argument("--pct-complete-override", type=float,
                    help="Use this % complete instead of EV/BAC (as a fraction)")
    ap.add_argument("--col", action="append", metavar="FIELD=COL",
                    help="Pin a column, e.g. --col ev=T --col ac=U. Repeatable. "
                         "Fields: code, desc, bac, pv, ev, ac, ffc, owner.")
    ap.add_argument("--json", help="Also write the full result to this JSON path")
    args = ap.parse_args()

    overrides = {}
    for item in (args.col or []):
        if "=" not in item:
            raise SystemExit(f"--col expects field=column, got '{item}'")
        k, _, v = item.partition("=")
        overrides[k.strip().lower()] = v.strip()

    o = analyse(args.data, args.sheet, args.header_row,
                args.pct_complete_override, overrides)
    print(render(o))
    if args.json:
        with open(args.json, "w", encoding="utf-8") as fh:
            json.dump(o, fh, indent=2, default=str)
        print(f"\nJSON written: {args.json}")


if __name__ == "__main__":
    sys.exit(main())
