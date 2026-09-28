#!/usr/bin/env python3
"""Build a complete, formula-live earned value workbook.

Sheets produced:
  Dashboard              headline indices, forecast, S-curve and CPI/SPI charts
  WPM                    Work Progress Measurement - where progress is entered
  CPR                    Cost Performance Report - rolls up by cost code
  Time-Phased Budget     period PV/EV/AC feeding the curves
  Variance Analysis      threshold breaches with cause / action / owner
  Budget Transfer Log    controlled baseline changes, feeds CPR variations column

Every figure is a live formula. Blue font marks cells the user fills in.

Usage:
    python build_evm_workbook.py --project "Name" --out out.xlsx [--data codes.csv]
                                 [--periods 12] [--threshold 0.10] [--currency AUD]

CSV columns (all optional, any subset works):
    cost_code, description, wbs_l1, wbs_l2, responsibility,
    budget_qty, unit, budget_value, technique, p6_id
"""

import argparse
import csv
import sys

from openpyxl import Workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

FONT = "Arial"

INPUT_FONT = Font(name=FONT, size=10, color="0000FF")
CALC_FONT = Font(name=FONT, size=10)
BOLD = Font(name=FONT, size=10, bold=True)
TITLE = Font(name=FONT, size=14, bold=True)
SUB = Font(name=FONT, size=10, italic=True, color="595959")
HDR_FONT = Font(name=FONT, size=9, bold=True, color="FFFFFF")
GRP_FONT = Font(name=FONT, size=10, bold=True, color="FFFFFF")

HDR_FILL = PatternFill("solid", fgColor="1F3864")
GRP_FILL = PatternFill("solid", fgColor="2E75B6")
TOT_FILL = PatternFill("solid", fgColor="D9E2F3")
IN_FILL = PatternFill("solid", fgColor="FFF2CC")
NOTE_FILL = PatternFill("solid", fgColor="F2F2F2")

THIN = Side(style="thin", color="BFBFBF")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

MONEY = '$#,##0;($#,##0);"-"'
MONEY2 = '$#,##0.00;($#,##0.00);"-"'
PCT = '0.0%;(0.0%);"-"'
IDX = '0.00'
QTY = '#,##0.00;(#,##0.00);"-"'

TECHNIQUES = "Units Complete,Milestone Weighting,Fixed Formula,% Complete,Level of Effort"

EXAMPLE_ROWS = [
    ("IND", "IND-TRC", "IND-TRC-01", "Site management staff", "Project Manager",
     18, "mth", 540000, "Level of Effort", "A1000"),
    ("DIR", "DIR-Z01", "DIR-Z01-EW", "Bulk earthworks", "Construction Manager",
     24000, "m3", 960000, "Units Complete", "EW1200"),
    ("DIR", "DIR-Z01", "DIR-Z01-CONC", "Concrete structures - substructure", "Construction Manager",
     1450, "m3", 1885000, "Units Complete", "CS1010"),
    ("DIR", "DIR-Z01", "DIR-Z01-STL", "Structural steel supply & install", "Construction Manager",
     380, "t", 1330000, "Milestone Weighting", "SS1018"),
    ("DIR", "DIR-Z01", "DIR-Z01-ELEC", "Electrical services", "Services Lead",
     1, "item", 720000, "Milestone Weighting", "EL1044"),
    ("DIR", "DIR-Z01", "DIR-Z01-TC", "Testing & commissioning", "Commissioning Manager",
     1, "item", 265000, "Milestone Weighting", "CM1070"),
]


def style_header(ws, row, first_col, last_col, height=34):
    ws.row_dimensions[row].height = height
    for c in range(first_col, last_col + 1):
        cell = ws.cell(row, c)
        cell.font = HDR_FONT
        cell.fill = HDR_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BOX


def group_band(ws, row, first_col, last_col, text):
    ws.cell(row, first_col, text)
    ws.merge_cells(start_row=row, start_column=first_col, end_row=row, end_column=last_col)
    for c in range(first_col, last_col + 1):
        cell = ws.cell(row, c)
        cell.fill = GRP_FILL
        cell.border = BOX
    cell = ws.cell(row, first_col)
    cell.font = GRP_FONT
    cell.alignment = Alignment(horizontal="center", vertical="center")


def titles(ws, project, subtitle, span):
    ws.cell(1, 1, project).font = TITLE
    ws.cell(2, 1, subtitle).font = SUB
    ws.cell(2, span, "Data date:").font = BOLD
    ws.cell(2, span).alignment = Alignment(horizontal="right")
    c = ws.cell(2, span + 1)
    c.font = INPUT_FONT
    c.fill = IN_FILL
    c.number_format = "dd-mmm-yy"
    c.border = BOX


def widths(ws, spec):
    for col, w in spec.items():
        ws.column_dimensions[col].width = w


def load_rows(path):
    if not path:
        return EXAMPLE_ROWS, True
    out = []
    with open(path, newline="", encoding="utf-8-sig") as fh:
        for r in csv.DictReader(fh):
            r = {(k or "").strip().lower(): (v or "").strip() for k, v in r.items()}
            # Skip rows that are entirely empty - exports routinely carry them, and a
            # blank row inside the block invites someone to type into it.
            if not any(r.values()):
                continue

            def num(key):
                try:
                    return float(r.get(key, "").replace(",", "").replace("$", ""))
                except (ValueError, AttributeError):
                    return None

            out.append((
                r.get("wbs_l1", ""), r.get("wbs_l2", ""), r.get("cost_code", ""),
                r.get("description", ""), r.get("responsibility", ""),
                num("budget_qty"), r.get("unit", ""), num("budget_value"),
                r.get("technique", ""), r.get("p6_id", ""),
            ))
    if not out:
        raise SystemExit("No rows found in --data file.")
    return out, False


# --------------------------------------------------------------------------- WPM

def build_wpm(wb, project, rows, example):
    ws = wb.create_sheet("WPM")
    titles(ws, project, "Work Progress Measurement - enter progress in the shaded columns each month", 17)

    group_band(ws, 4, 1, 5, "SCOPE")
    group_band(ws, 4, 6, 10, "PERFORMANCE MEASUREMENT BASELINE")
    group_band(ws, 4, 11, 17, "PROGRESS THIS PERIOD")
    group_band(ws, 4, 18, 18, "")

    hdr = ["WBS L1", "WBS L2", "Cost Code", "Detail Description", "Responsibility",
           "Budget\nQuantity", "Unit", "Budget Value", "EV Rate\n/Unit", "Performance\nTechnique",
           "Quantity\nComplete\nto Date", "Quantity\nRemaining", "Quantity At\nCompletion",
           "Variance\nBudget Qty v\nQty at Compl.", "Earned Value", "%\nComplete", "Status",
           "P6 Activity ID"]
    for i, h in enumerate(hdr, start=1):
        ws.cell(5, i, h)
    style_header(ws, 5, 1, 18, height=44)

    first = 6
    last = first + len(rows) - 1

    for i, r in enumerate(rows):
        rw = first + i
        wbs1, wbs2, code, desc, resp, qty, unit, val, tech, p6 = r
        for col, v in ((1, wbs1), (2, wbs2), (3, code), (4, desc), (5, resp),
                       (6, qty), (7, unit), (8, val), (10, tech), (18, p6)):
            c = ws.cell(rw, col, v)
            c.font = INPUT_FONT
            c.fill = IN_FILL
        ws.cell(rw, 6).number_format = QTY
        ws.cell(rw, 8).number_format = MONEY

        # Progress inputs - left blank for the user, including in the example set.
        for col in (11, 12):
            c = ws.cell(rw, col)
            c.font = INPUT_FONT
            c.fill = IN_FILL
            c.number_format = QTY

        ws.cell(rw, 9, f"=IF(F{rw}=0,0,H{rw}/F{rw})").number_format = MONEY2
        ws.cell(rw, 13, f"=K{rw}+L{rw}").number_format = QTY
        ws.cell(rw, 14, f"=F{rw}-M{rw}").number_format = QTY
        ws.cell(rw, 15, f"=IF(M{rw}=0,0,(K{rw}/M{rw})*H{rw})").number_format = MONEY
        ws.cell(rw, 16, f"=IF(H{rw}=0,0,O{rw}/H{rw})").number_format = PCT
        ws.cell(rw, 17,
                f'=IF(ROUND(P{rw},2)=1,"Completed",'
                f'IF(ROUND(P{rw},2)>0,"In Progress","Not Started"))')
        for col in range(1, 19):
            c = ws.cell(rw, col)
            c.border = BOX
            if c.font.color is None or c.font.color.rgb != "000000FF":
                if col not in (1, 2, 3, 4, 5, 6, 7, 8, 10, 11, 12, 18):
                    c.font = CALC_FONT

    tot = last + 1
    ws.cell(tot, 4, "PROJECT TOTAL - Performance Measurement Baseline").font = BOLD
    ws.cell(tot, 8, f"=SUBTOTAL(9,H{first}:H{last})").number_format = MONEY
    ws.cell(tot, 15, f"=SUBTOTAL(9,O{first}:O{last})").number_format = MONEY
    ws.cell(tot, 16, f"=IF(H{tot}=0,0,O{tot}/H{tot})").number_format = PCT
    for c in range(1, 19):
        cell = ws.cell(tot, c)
        cell.fill = TOT_FILL
        cell.border = BOX
        if cell.font.b is not True:
            cell.font = BOLD

    dv = DataValidation(type="list", formula1=f'"{TECHNIQUES}"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(f"J{first}:J{last}")

    note = tot + 2
    ws.cell(note, 1, "Blue cells are inputs. Enter quantity complete and quantity remaining "
                     "each month; earned value, % complete and status calculate themselves.").font = SUB
    ws.cell(note + 1, 1, "Pick the performance technique before work starts and do not change it "
                         "mid-package. Never record earned value against allowances or contingency.").font = SUB
    if example:
        ws.cell(note + 2, 1, "The rows above are a worked example showing the expected format - "
                             "replace them with the real cost codes.").font = SUB

    widths(ws, {"A": 9, "B": 9, "C": 12, "D": 42, "E": 20, "F": 12, "G": 8, "H": 14,
                "I": 12, "J": 18, "K": 12, "L": 12, "M": 12, "N": 13, "O": 14, "P": 10,
                "Q": 12, "R": 14})
    ws.freeze_panes = "E6"
    return first, last, tot


# --------------------------------------------------------------------------- CPR

def build_cpr(wb, project, rows, wpm_first, wpm_last, threshold):
    ws = wb.create_sheet("CPR")
    titles(ws, project, "Cost Performance Report - cumulative to date", 24)

    ws.cell(3, 1, "Variance threshold:").font = BOLD
    t = ws.cell(3, 2, threshold)
    t.font = INPUT_FONT
    t.fill = IN_FILL
    t.number_format = "0%"
    t.border = BOX
    ws.cell(3, 3, "Cost codes breaching this on cost or schedule require commentary "
                  "and a corrective action.").font = SUB

    group_band(ws, 5, 1, 2, "COST ACCOUNT")
    group_band(ws, 5, 3, 7, "BUDGET")
    group_band(ws, 5, 8, 10, "CUMULATIVE TO DATE")
    group_band(ws, 5, 11, 16, "VARIANCE & PERFORMANCE")
    group_band(ws, 5, 17, 20, "AT COMPLETION")
    group_band(ws, 5, 21, 23, "INDEPENDENT CHECK")
    group_band(ws, 5, 24, 27, "ACCOUNTABILITY")

    hdr = ["Cost Code", "Description",
           "Original\nBudget", "Variations &\nTransfers", "Budget at\nCompletion",
           "%\nComplete", "%\nSpent",
           "Planned\nValue", "Earned\nValue", "Actual\nCost",
           "CV $", "CV %", "SV $", "SV %", "CPI", "SPI",
           "Forecast\nFinal Cost", "Estimate to\nComplete", "VAC $", "VAC %",
           "IEAC", "IEAC\nVariance", "TCPI\n(to BAC)",
           "Data Check", "Status", "Responsibility", "Comments"]
    for i, h in enumerate(hdr, start=1):
        ws.cell(6, i, h)
    style_header(ws, 6, 1, 27, height=46)

    codes = []
    seen = set()
    for r in rows:
        code = r[2]
        if code and code not in seen:
            seen.add(code)
            codes.append((code, r[3], r[4]))
    if not codes:
        codes = [("", "", "")]

    first = 7
    last = first + len(codes) - 1
    wpm_code = f"WPM!$C${wpm_first}:$C${wpm_last}"

    for i, (code, desc, resp) in enumerate(codes):
        rw = first + i
        c = ws.cell(rw, 1, code)
        c.font = INPUT_FONT
        c.fill = IN_FILL
        ws.cell(rw, 2, f'=IFERROR(INDEX(WPM!$D${wpm_first}:$D${wpm_last},'
                       f'MATCH(A{rw},{wpm_code},0)),"")')
        ws.cell(rw, 3, f"=SUMIF({wpm_code},A{rw},WPM!$H${wpm_first}:$H${wpm_last})").number_format = MONEY
        ws.cell(rw, 4, f"=SUMIF('Budget Transfer Log'!$D:$D,A{rw},'Budget Transfer Log'!$F:$F)"
                       f"-SUMIF('Budget Transfer Log'!$C:$C,A{rw},'Budget Transfer Log'!$F:$F)"
                       ).number_format = MONEY
        ws.cell(rw, 5, f"=C{rw}+D{rw}").number_format = MONEY
        ws.cell(rw, 6, f"=IF(E{rw}=0,0,I{rw}/E{rw})").number_format = PCT
        ws.cell(rw, 7, f"=IF(E{rw}=0,0,J{rw}/E{rw})").number_format = PCT

        pv = ws.cell(rw, 8)          # planned value - from the programme
        pv.font = INPUT_FONT
        pv.fill = IN_FILL
        pv.number_format = MONEY

        ws.cell(rw, 9, f"=SUMIF({wpm_code},A{rw},WPM!$O${wpm_first}:$O${wpm_last})").number_format = MONEY

        ac = ws.cell(rw, 10)         # actual cost - from the cost ledger, incl. accruals
        ac.font = INPUT_FONT
        ac.fill = IN_FILL
        ac.number_format = MONEY

        ws.cell(rw, 11, f"=I{rw}-J{rw}").number_format = MONEY
        ws.cell(rw, 12, f"=IF(I{rw}=0,0,K{rw}/I{rw})").number_format = PCT
        ws.cell(rw, 13, f"=I{rw}-H{rw}").number_format = MONEY
        ws.cell(rw, 14, f"=IF(H{rw}=0,0,M{rw}/H{rw})").number_format = PCT
        ws.cell(rw, 15, f'=IF(J{rw}=0,"",I{rw}/J{rw})').number_format = IDX
        ws.cell(rw, 16, f'=IF(H{rw}=0,"",I{rw}/H{rw})').number_format = IDX

        ffc = ws.cell(rw, 17, f"=E{rw}")   # default to BAC until the team forecasts
        ffc.font = INPUT_FONT
        ffc.fill = IN_FILL
        ffc.number_format = MONEY

        ws.cell(rw, 18, f"=Q{rw}-J{rw}").number_format = MONEY
        ws.cell(rw, 19, f"=E{rw}-Q{rw}").number_format = MONEY
        ws.cell(rw, 20, f'=IF(E{rw}=0,"",S{rw}/E{rw})').number_format = PCT
        # IEAC suppressed below 30% complete - CPI has not stabilised that early.
        ws.cell(rw, 21, f'=IF(OR(F{rw}<0.3,O{rw}="",O{rw}=0),"n/a",E{rw}/O{rw})').number_format = MONEY
        ws.cell(rw, 22, f'=IF(ISNUMBER(U{rw}),Q{rw}-U{rw},"")').number_format = MONEY
        ws.cell(rw, 23, f'=IF((E{rw}-J{rw})<=0,"budget spent",(E{rw}-I{rw})/(E{rw}-J{rw}))'
                ).number_format = IDX
        # Integrity checks run before the variance flag: a threshold breach on bad
        # data sends people chasing a cause that does not exist.
        ws.cell(rw, 24,
                f'=IF(AND(I{rw}>0,J{rw}=0),"EV no AC",'
                f'IF(AND(J{rw}>0,I{rw}=0),"AC no EV",'
                f'IF(I{rw}>E{rw}*1.0001,"EV over budget",'
                f'IF(AND(E{rw}>0,Q{rw}<J{rw}),"FFC below AC",'
                f'IF(AND(ROUND(F{rw},4)>=1,R{rw}>0),"Done, ETC left",'
                f'IF(AND(H{rw}>0,I{rw}=0,J{rw}=0),"PV no progress",""))))))')
        ws.cell(rw, 25, f'=IF(AND(I{rw}=0,J{rw}=0),"Not started",'
                        f'IF(X{rw}<>"","CHECK DATA",'
                        f'IF(OR(ABS(L{rw})>$B$3,ABS(N{rw})>$B$3),"REVIEW","OK")))')
        ws.cell(rw, 26, f'=IFERROR(INDEX(WPM!$E${wpm_first}:$E${wpm_last},'
                        f'MATCH(A{rw},{wpm_code},0)),"")')
        cm = ws.cell(rw, 27)
        cm.font = INPUT_FONT
        cm.fill = IN_FILL

        for col in range(1, 28):
            cell = ws.cell(rw, col)
            cell.border = BOX
            if col not in (1, 8, 10, 17, 27):
                cell.font = CALC_FONT

    tot = last + 1
    ws.cell(tot, 2, "PROJECT TOTAL").font = BOLD
    for col in (3, 4, 5, 8, 9, 10, 17, 18, 19):
        L = get_column_letter(col)
        ws.cell(tot, col, f"=SUM({L}{first}:{L}{last})").number_format = MONEY
    ws.cell(tot, 6, f"=IF(E{tot}=0,0,I{tot}/E{tot})").number_format = PCT
    ws.cell(tot, 7, f"=IF(E{tot}=0,0,J{tot}/E{tot})").number_format = PCT
    ws.cell(tot, 11, f"=I{tot}-J{tot}").number_format = MONEY
    ws.cell(tot, 12, f"=IF(I{tot}=0,0,K{tot}/I{tot})").number_format = PCT
    ws.cell(tot, 13, f"=I{tot}-H{tot}").number_format = MONEY
    ws.cell(tot, 14, f"=IF(H{tot}=0,0,M{tot}/H{tot})").number_format = PCT
    ws.cell(tot, 15, f'=IF(J{tot}=0,"",I{tot}/J{tot})').number_format = IDX
    ws.cell(tot, 16, f'=IF(H{tot}=0,"",I{tot}/H{tot})').number_format = IDX
    ws.cell(tot, 20, f'=IF(E{tot}=0,"",S{tot}/E{tot})').number_format = PCT
    ws.cell(tot, 21, f'=IF(OR(F{tot}<0.3,O{tot}="",O{tot}=0),"n/a",E{tot}/O{tot})').number_format = MONEY
    ws.cell(tot, 22, f'=IF(ISNUMBER(U{tot}),Q{tot}-U{tot},"")').number_format = MONEY
    ws.cell(tot, 23, f'=IF((E{tot}-J{tot})<=0,"budget spent",(E{tot}-I{tot})/(E{tot}-J{tot}))'
            ).number_format = IDX
    ws.cell(tot, 24, f'=COUNTIF(X{first}:X{last},"?*")&" flagged"')
    ws.cell(tot, 25, f'=IF(AND(I{tot}=0,J{tot}=0),"Not started",'
                     f'IF(OR(ABS(L{tot})>$B$3,ABS(N{tot})>$B$3),"REVIEW","OK"))')
    for c in range(1, 28):
        cell = ws.cell(tot, c)
        cell.fill = TOT_FILL
        cell.border = BOX
        cell.font = BOLD

    note = tot + 2
    ws.cell(note, 1, "Blue cells are inputs: planned value from the programme, actual cost from "
                     "the ledger including accruals, and the team's forecast final cost.").font = SUB
    ws.cell(note + 1, 1, "Original budget and earned value pull from WPM by cost code. Variations "
                         "pull from the Budget Transfer Log. IEAC is suppressed below 30% "
                         "complete because CPI has not stabilised.").font = SUB
    ws.cell(note + 2, 1, "Keep PV, EV and AC on the same basis - all excluding GST, or all "
                         "including. Mixing them is the most common error in a cost report.").font = SUB

    widths(ws, {"A": 12, "B": 34, "C": 14, "D": 13, "E": 14, "F": 10, "G": 10, "H": 14,
                "I": 14, "J": 14, "K": 13, "L": 9, "M": 13, "N": 9, "O": 8, "P": 8,
                "Q": 14, "R": 14, "S": 13, "T": 9, "U": 14, "V": 13, "W": 10, "X": 14,
                "Y": 12, "Z": 18, "AA": 40})
    ws.freeze_panes = "C7"
    return first, last, tot


# ------------------------------------------------------------------ time-phased

def build_timephased(wb, project, periods, cpr_tot):
    ws = wb.create_sheet("Time-Phased Budget")
    titles(ws, project, "Time-phased budget and actuals - drives the S-curve and index trend", 8)

    hdr = ["Period", "PV\nthis period", "EV\nthis period", "AC\nthis period",
           "Cumulative\nPV", "Cumulative\nEV", "Cumulative\nAC", "CPI", "SPI"]
    for i, h in enumerate(hdr, start=1):
        ws.cell(5, i, h)
    style_header(ws, 5, 1, 9)

    first = 6
    last = first + periods - 1
    for i in range(periods):
        rw = first + i
        p = ws.cell(rw, 1, f"P{i + 1}")
        p.font = INPUT_FONT
        p.fill = IN_FILL
        for col in (2, 3, 4):
            c = ws.cell(rw, col)
            c.font = INPUT_FONT
            c.fill = IN_FILL
            c.number_format = MONEY
        for src, dst in ((2, 5), (3, 6), (4, 7)):
            S = get_column_letter(src)
            ws.cell(rw, dst, f"=SUM(${S}${first}:{S}{rw})").number_format = MONEY
        ws.cell(rw, 8, f'=IF(G{rw}=0,"",F{rw}/G{rw})').number_format = IDX
        ws.cell(rw, 9, f'=IF(E{rw}=0,"",F{rw}/E{rw})').number_format = IDX
        for col in range(1, 10):
            cell = ws.cell(rw, col)
            cell.border = BOX
            if col > 4:
                cell.font = CALC_FONT

    note = last + 2
    ws.cell(note, 1, "Enter the planned value profile for every period up front - that curve is "
                     "the baseline. Add earned value and actual cost as each month closes; "
                     "leave future periods blank so the curves stop at the data date.").font = SUB
    ws.cell(note + 1, 1, f"Cumulative EV at the final period should reconcile to the CPR total "
                         f"earned value (CPR row {cpr_tot}).").font = SUB

    widths(ws, {"A": 12, "B": 15, "C": 15, "D": 15, "E": 15, "F": 15, "G": 15, "H": 9, "I": 9})
    ws.freeze_panes = "B6"
    return first, last


def add_charts(wb, ws, tp_first, tp_last):
    tp = "Time-Phased Budget"
    s = LineChart()
    s.title = "Progress S-Curve - cumulative PV, EV and AC"
    s.style = 2
    s.y_axis.title = "Cumulative value"
    s.x_axis.title = "Period"
    s.height, s.width = 9.5, 19
    data = Reference(wb[tp], min_col=5, max_col=7, min_row=5, max_row=tp_last)
    cats = Reference(wb[tp], min_col=1, min_row=tp_first, max_row=tp_last)
    s.add_data(data, titles_from_data=True)
    s.set_categories(cats)
    ws.add_chart(s, "A22")

    i = LineChart()
    i.title = "CPI and SPI trend"
    i.style = 2
    i.y_axis.title = "Index"
    i.x_axis.title = "Period"
    i.height, i.width = 9.5, 19
    idata = Reference(wb[tp], min_col=8, max_col=9, min_row=5, max_row=tp_last)
    i.add_data(idata, titles_from_data=True)
    i.set_categories(cats)
    ws.add_chart(i, "A44")


# ------------------------------------------------------------ variance analysis

def build_var(wb, project, cpr_first, cpr_last, threshold):
    ws = wb.create_sheet("Variance Analysis")
    titles(ws, project, "Variance Analysis Report - every cost code breaching the threshold "
                        "needs a cause, an action and an owner", 9)

    hdr = ["Cost Code", "Description", "CV $", "CV %", "SV $", "SV %",
           "Data Check", "Status",
           "Cause", "One-off or systemic?", "Impact on forecast",
           "Corrective action", "Owner", "Due"]
    for i, h in enumerate(hdr, start=1):
        ws.cell(5, i, h)
    style_header(ws, 5, 1, 14)

    first = 6
    for i in range(cpr_last - cpr_first + 1):
        rw = first + i
        src = cpr_first + i
        ws.cell(rw, 1, f"=CPR!A{src}")
        ws.cell(rw, 2, f"=CPR!B{src}")
        ws.cell(rw, 3, f"=CPR!K{src}").number_format = MONEY
        ws.cell(rw, 4, f"=CPR!L{src}").number_format = PCT
        ws.cell(rw, 5, f"=CPR!M{src}").number_format = MONEY
        ws.cell(rw, 6, f"=CPR!N{src}").number_format = PCT
        ws.cell(rw, 7, f"=CPR!X{src}")
        ws.cell(rw, 8, f"=CPR!Y{src}")
        for col in range(9, 15):
            c = ws.cell(rw, col)
            c.font = INPUT_FONT
            c.fill = IN_FILL
        ws.cell(rw, 14).number_format = "dd-mmm-yy"
        for col in range(1, 15):
            cell = ws.cell(rw, col)
            cell.border = BOX
            if col < 9:
                cell.font = CALC_FONT

    note = first + (cpr_last - cpr_first) + 2
    ws.cell(note, 1, f"Clear every Data Check flag before writing commentary - a threshold "
                     f"breach on bad data sends people chasing a cause that is not there. "
                     f"Then filter Status to REVIEW "
                     f"(threshold {threshold:.0%}, set on the CPR sheet).").font = SUB
    ws.cell(note + 1, 1, "One-off or systemic is the load-bearing call: it decides which EAC "
                         "formula is honest. One-off means the remaining work runs to budget; "
                         "systemic means current performance continues.").font = SUB
    ws.cell(note + 2, 1, "Positive variances need explaining too - a favourable CV on a lump-sum "
                         "package is usually buying gain banked at award, not efficiency, and it "
                         "will not repeat.").font = SUB

    widths(ws, {"A": 12, "B": 32, "C": 13, "D": 9, "E": 13, "F": 9, "G": 14, "H": 12,
                "I": 38, "J": 20, "K": 30, "L": 38, "M": 18, "N": 12})
    ws.freeze_panes = "C6"


# ------------------------------------------------------------------ transfer log

def build_transfer_log(wb, project):
    ws = wb.create_sheet("Budget Transfer Log")
    titles(ws, project, "Budget transfers and variations - the only sanctioned route to change "
                        "the baseline", 7)

    hdr = ["Date", "Reference", "From Cost Code", "To Cost Code", "Type",
           "Amount", "Reason", "Approved by"]
    for i, h in enumerate(hdr, start=1):
        ws.cell(5, i, h)
    style_header(ws, 5, 1, 8)

    for rw in range(6, 46):
        for col in range(1, 9):
            c = ws.cell(rw, col)
            c.font = INPUT_FONT
            c.fill = IN_FILL
            c.border = BOX
        ws.cell(rw, 1).number_format = "dd-mmm-yy"
        ws.cell(rw, 6).number_format = MONEY

    dv = DataValidation(
        type="list",
        formula1='"Client Variation,Design Development,Risk & Opportunity,Escalation,'
                 'Methodology Change,Re-baseline"',
        allow_blank=True)
    ws.add_data_validation(dv)
    dv.add("E6:E45")

    ws.cell(47, 1, "The CPR variations column reads this log: amounts to a cost code add, "
                   "amounts from it subtract. Close transfers two working days before month end.").font = SUB
    ws.cell(48, 1, "Budget for completed work is never changed, and contingency is never used to "
                   "absorb an overrun - that hides the variance the system exists to expose.").font = SUB

    widths(ws, {"A": 12, "B": 16, "C": 16, "D": 16, "E": 22, "F": 15, "G": 46, "H": 18})
    ws.freeze_panes = "A6"


# -------------------------------------------------------------------- dashboard

def build_dashboard(wb, project, cpr_tot, threshold, currency):
    ws = wb["Dashboard"]
    ws.cell(1, 1, project).font = Font(name=FONT, size=18, bold=True)
    ws.cell(2, 1, f"Earned value dashboard - all figures {currency}, "
                  f"variance threshold {threshold:.0%}").font = SUB

    metrics = [
        ("Budget at Completion (BAC)", f"=CPR!E{cpr_tot}", MONEY),
        ("Planned Value (PV)", f"=CPR!H{cpr_tot}", MONEY),
        ("Earned Value (EV)", f"=CPR!I{cpr_tot}", MONEY),
        ("Actual Cost (AC)", f"=CPR!J{cpr_tot}", MONEY),
        ("% Complete", f"=CPR!F{cpr_tot}", PCT),
        ("% Spent", f"=CPR!G{cpr_tot}", PCT),
        ("Cost Variance (CV)", f"=CPR!K{cpr_tot}", MONEY),
        ("Schedule Variance (SV)", f"=CPR!M{cpr_tot}", MONEY),
        ("CPI", f"=CPR!O{cpr_tot}", IDX),
        ("SPI", f"=CPR!P{cpr_tot}", IDX),
        ("Forecast Final Cost (FFC)", f"=CPR!Q{cpr_tot}", MONEY),
        ("Independent EAC", f"=CPR!U{cpr_tot}", MONEY),
        ("FFC v IEAC gap", f"=CPR!V{cpr_tot}", MONEY),
        ("Variance at Completion (VAC)", f"=CPR!S{cpr_tot}", MONEY),
        ("TCPI to BAC", f"=CPR!W{cpr_tot}", IDX),
    ]
    ws.cell(4, 1, "HEADLINE").font = BOLD
    for i, (label, formula, fmt) in enumerate(metrics):
        rw = 5 + i
        lc = ws.cell(rw, 1, label)
        lc.font = CALC_FONT
        lc.border = BOX
        vc = ws.cell(rw, 2, formula)
        vc.font = BOLD
        vc.number_format = fmt
        vc.border = BOX
        vc.alignment = Alignment(horizontal="right")

    # Verdicts are gated on project maturity and use a tolerance band, because a
    # raw CPI<1 test calls "over budget" on a job at 2% complete with a rounding
    # difference, and that is how a cost report loses its audience.
    t = cpr_tot
    pct = f"CPR!F{t}"        # % complete
    cpi = f"CPR!O{t}"
    spi = f"CPR!P{t}"
    tcpi = f"CPR!W{t}"
    started = f'AND(ISNUMBER({cpi}),CPR!J{t}>0)'
    mature = f"{pct}>=0.15"   # below this, indices are noise
    # The tolerance band follows the project's own variance threshold, so the
    # dashboard and the CPR flags never disagree about what counts as on budget.
    lo = f"(1-CPR!$B$3)"
    hi = f"(1+CPR!$B$3)"

    ws.cell(4, 4, "READING THIS").font = BOLD
    reads = [
        ("Cost",
         f'=IF(NOT({started}),"Not started - no cost performance yet",'
         f'IF(NOT({mature}),"Below 15% complete - CPI "&TEXT({cpi},"0.00")&'
         f'" is not yet a reliable signal",'
         f'IF({cpi}<{lo},"Over budget - CPI "&TEXT({cpi},"0.00"),'
         f'IF({cpi}<={hi},"On budget within tolerance - CPI "&TEXT({cpi},"0.00"),'
         f'"Favourable - CPI "&TEXT({cpi},"0.00")&'
         f'". Check this is real efficiency, not buying gain or unposted cost"))))'),
        ("Schedule",
         f'=IF(OR(NOT(ISNUMBER({spi})),CPR!I{t}=0),"Not started - no schedule performance yet",'
         f'IF(NOT({mature}),"Below 15% complete - SPI swings wildly on small numbers",'
         f'IF({pct}>0.7,"Past 70% complete - SPI drifts to 1.0 regardless of lateness. '
         f'Read the programme, not SPI.",'
         f'IF({spi}<{lo},"Behind - SPI "&TEXT({spi},"0.00"),'
         f'IF({spi}<={hi},"On schedule within tolerance",'
         f'"Ahead - SPI "&TEXT({spi},"0.00"))))))'),
        ("Recovery",
         f'=IF(NOT({started}),"Nothing to recover yet",'
         f'IF(NOT({mature}),"Below 15% complete - TCPI not yet meaningful",'
         f'IF(ROUND({pct},4)>=1,"Complete - final CPI "&TEXT({cpi},"0.00")&'
         f'", nothing left to recover",'
         f'IF(NOT(ISNUMBER({tcpi})),"Budget already spent with work remaining - reset the target",'
         f'IF({pct}>0.9,"Past 90% complete - TCPI is unstable on the small remainder",'
         f'IF({tcpi}>1.1,"Not credible - TCPI "&TEXT({tcpi},"0.00")&'
         f'" against CPI "&TEXT({cpi},"0.00")&". Reset the target.",'
         f'IF({tcpi}>1.02,"Recovery required - TCPI "&TEXT({tcpi},"0.00"),'
         f'"Target consistent with performance")))))))'),
        ("Forecast",
         f'=IF(NOT(ISNUMBER(CPR!U{t})),"IEAC suppressed below 30% complete - CPI has not settled",'
         f'IF(CPR!Q{t}<CPR!U{t}*0.98,"FFC is "&TEXT(CPR!U{t}-CPR!Q{t},"$#,##0")&'
         f'" below the independent estimate - what specifically changes to close that?",'
         f'IF(CPR!Q{t}>CPR!U{t}*1.02,"FFC sits above the independent estimate - conservative",'
         f'"FFC consistent with the independent estimate")))'),
        ("Data check",
         f'=IF(AND(CPR!I{t}>0,CPR!J{t}=0),"Earned value with no actual cost - costs unposted or miscoded",'
         f'IF(AND(CPR!J{t}>0,CPR!I{t}=0),"Actual cost with no earned value - progress under-claimed",'
         f'IF(CPR!I{t}>CPR!E{t}*1.0001,"Earned value exceeds budget - measurement error",'
         f'IF(CPR!Q{t}<CPR!J{t},"Forecast is below cost already spent - impossible",'
         f'"No headline data integrity flags"))))'),
    ]
    for i, (label, formula) in enumerate(reads):
        rw = 5 + i
        lc = ws.cell(rw, 4, label)
        lc.font = CALC_FONT
        lc.border = BOX
        vc = ws.cell(rw, 5, formula)
        vc.font = CALC_FONT
        vc.border = BOX

    ws.cell(11, 4, "WHERE THE NUMBERS COME FROM").font = BOLD
    src = [
        "WPM - enter quantity complete and remaining; earned value calculates.",
        "CPR - enter planned value from the programme, actual cost from the ledger, "
        "and the team forecast.",
        "Time-Phased Budget - enter the PV profile once, then EV and AC each period.",
        "Budget Transfer Log - every approved variation and transfer.",
        "Variance Analysis - explain each REVIEW flag with a cause, action and owner.",
    ]
    for i, line in enumerate(src):
        c = ws.cell(12 + i, 4, line)
        c.font = SUB

    widths(ws, {"A": 32, "B": 18, "C": 3, "D": 22, "E": 78})


# ------------------------------------------------------------------------- main

def add_defined_names(wb, wpm_first, wpm_last, wpm_tot,
                      cpr_first, cpr_last, cpr_tot, tp_first, tp_last):
    """Name the data blocks and total rows.

    Without these, nothing downstream - a later script, a BI tool, or a person
    writing a formula in another workbook - can find where the data stops, because
    every cell in the block holds a formula and the total row looks like any other.
    """
    names = {
        "WPM_DATA": f"WPM!$A${wpm_first}:$R${wpm_last}",
        "WPM_TOTAL": f"WPM!$A${wpm_tot}:$R${wpm_tot}",
        "WPM_EV": f"WPM!$O${wpm_first}:$O${wpm_last}",
        "CPR_DATA": f"CPR!$A${cpr_first}:$AA${cpr_last}",
        "CPR_TOTAL": f"CPR!$A${cpr_tot}:$AA${cpr_tot}",
        "PMB_BAC": f"CPR!$E${cpr_tot}",
        "PMB_PV": f"CPR!$H${cpr_tot}",
        "PMB_EV": f"CPR!$I${cpr_tot}",
        "PMB_AC": f"CPR!$J${cpr_tot}",
        "TP_DATA": f"'Time-Phased Budget'!$A${tp_first}:$I${tp_last}",
    }
    for name, ref in names.items():
        wb.defined_names.add(DefinedName(name, attr_text=ref))


def main():
    ap = argparse.ArgumentParser(description="Build an earned value workbook.")
    ap.add_argument("--project", required=True, help="Project name shown on every sheet")
    ap.add_argument("--out", required=True, help="Output .xlsx path")
    ap.add_argument("--data", help="CSV of cost codes (see module docstring)")
    ap.add_argument("--periods", type=int, default=12, help="Reporting periods (default 12)")
    ap.add_argument("--threshold", type=float, default=0.10,
                    help="Variance threshold as a fraction (default 0.10)")
    ap.add_argument("--currency", default="AUD", help="Currency label (default AUD)")
    args = ap.parse_args()

    rows, example = load_rows(args.data)

    wb = Workbook()
    wb.active.title = "Dashboard"

    wpm_first, wpm_last, wpm_tot = build_wpm(wb, args.project, rows, example)
    cpr_first, cpr_last, cpr_tot = build_cpr(
        wb, args.project, rows, wpm_first, wpm_last, args.threshold)
    tp_first, tp_last = build_timephased(wb, args.project, args.periods, cpr_tot)
    build_var(wb, args.project, cpr_first, cpr_last, args.threshold)
    build_transfer_log(wb, args.project)
    build_dashboard(wb, args.project, cpr_tot, args.threshold, args.currency)
    add_charts(wb, wb["Dashboard"], tp_first, tp_last)
    add_defined_names(wb, wpm_first, wpm_last, wpm_tot,
                      cpr_first, cpr_last, cpr_tot, tp_first, tp_last)

    wb.save(args.out)
    print(f"Written: {args.out}")
    print(f"  cost codes: {len(rows)}   periods: {args.periods}   "
          f"threshold: {args.threshold:.0%}   example data: {example}")
    print("Now run the xlsx skill's recalc.py before delivering.")


if __name__ == "__main__":
    sys.exit(main())
