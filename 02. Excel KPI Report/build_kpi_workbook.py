"""
Build the formula-driven Excel KPI workbook from the cleaned CSVs.

Every KPI in the workbook is an Excel formula (COUNTIFS, SUMIFS, AVERAGEIFS, SUMPRODUCT, INDEX/MATCH), so the report
recalculates when the Jobs data is replaced with a new month's extract. Python only loads the data and lays out the sheets.

Run:  python build_kpi_workbook.py      (after 01. Dataset/clean_dataset.py)
Output: Operations_KPI_Report.xlsx  (open in Excel; formulas calculate on open)
"""
import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

HERE = os.path.dirname(os.path.abspath(__file__))
DATA = os.path.join(HERE, "..", "01. Dataset", "Cleaned Dataset")
OUT = os.path.join(HERE, "Operations_KPI_Report.xlsx")

F = "Arial"
H1 = Font(name=F, size=16, bold=True)
H2 = Font(name=F, size=11, bold=True, color="FFFFFF")
B = Font(name=F, size=10)
BB = Font(name=F, size=10, bold=True)
BLUE = Font(name=F, size=10, color="0000FF")
GREY = Font(name=F, size=9, italic=True, color="666666")
HEAD = PatternFill("solid", fgColor="2A78D6")
YEL = PatternFill("solid", fgColor="FFFF00")
LIGHT = PatternFill("solid", fgColor="F2F2F2")
thin = Side(style="thin", color="D9D9D9")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
GOOD, WARN, BAD = (PatternFill("solid", fgColor=c) for c in ("C6EFCE", "FFEB9C", "FFC7CE"))

jobs = pd.read_csv(os.path.join(DATA, "Clean_Jobs.csv"), keep_default_na=False)
staff = pd.read_csv(os.path.join(DATA, "Clean_Staffing.csv"))
esc = pd.read_csv(os.path.join(DATA, "Clean_Escalations.csv"))
months = sorted(staff.Month.unique())
teams = sorted(staff.Team.unique())
NJ, NS, NE = len(jobs) + 1, len(staff) + 1, len(esc) + 1     # last data row
wb = Workbook()


def header(ws, row, cols, widths=None):
    for i, c in enumerate(cols, start=1):
        cell = ws.cell(row=row, column=i, value=c)
        cell.font, cell.fill, cell.alignment, cell.border = H2, HEAD, Alignment(horizontal="center", vertical="center", wrap_text=True), BOX
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w


def rng_(col, last):
    return f"Jobs!${col}$2:${col}${last}"


J = {k: rng_(c, NJ) for k, c in zip(["id", "client", "team", "type", "rec", "due", "del", "pages", "err", "rework", "status",
                                      "recm", "delm", "tat", "ontime"], "ABCDEFGHIJKLMNO")}

# ------------------------------------------------------------------ Read Me
ws = wb.active
ws.title = "Read Me"
ws["A1"], ws["A1"].font = "Operations KPI Report - Content Services (synthetic data)", H1
rows = [
    ("Purpose", "Monthly operational KPI report for three production teams: Data Conversion, Academic Editing, Typesetting."),
    ("Data", "Synthetic jobs, staffing and escalations, Jul 2023 - Dec 2024 (see 01. Dataset). No real company data."),
    ("How to refresh", "Paste the new cleaned extract into Jobs (A:K), Staffing and Escalations; extend the helper formulas in Jobs L:O; add the new month to Monthly KPIs."),
    ("Monthly KPIs", "One row per month. Every number is a formula over the Jobs, Staffing and Escalations sheets."),
    ("Team Scorecard", "Pick a month in the yellow cell C3. Shows each team against targets and the previous month."),
    ("Targets", "KPI targets in blue are inputs: change them and the RAG colours update."),
    ("DQ Checklist", "Automated data-quality checks (formulas) plus the monthly manual checklist."),
    ("KPI Definitions", "Formula and meaning of every KPI."),
    ("Colour legend", "Blue text = input you may change. Yellow fill = selector. Black = formula. Green / amber / red fill = on target / within 5 pts / off target."),
]
for i, (a, b) in enumerate(rows, start=3):
    ws.cell(row=i, column=1, value=a).font = BB
    ws.cell(row=i, column=2, value=b).font = B
ws.column_dimensions["A"].width, ws.column_dimensions["B"].width = 18, 120

# ------------------------------------------------------------------ Targets
tg = wb.create_sheet("Targets")
tg["A1"], tg["A1"].font = "KPI targets (inputs)", H1
header(tg, 3, ["KPI", "Target", "Direction", "Note"], [32, 12, 14, 60])
targets = [("On-time delivery %", 0.90, "Higher is better", "Delivered on or before due date"),
           ("First-pass quality %", 0.95, "Higher is better", "Delivered without rework"),
           ("Average turnaround (days)", 5.0, "Lower is better", "Received to delivered"),
           ("Errors per 1,000 pages", 2.5, "Lower is better", "QA errors found / pages delivered x 1,000"),
           ("Staffing vs plan %", 0.95, "Higher is better", "Actual FTE / planned FTE")]
for i, (k, v, dr, n) in enumerate(targets, start=4):
    tg.cell(row=i, column=1, value=k).font = B
    c = tg.cell(row=i, column=2, value=v)
    c.font, c.number_format = BLUE, "0.0%" if isinstance(v, float) and v < 1.5 else "0.0"
    tg.cell(row=i, column=3, value=dr).font = B
    tg.cell(row=i, column=4, value=n).font = GREY
tg["A11"], tg["A11"].font = "Targets are illustrative and were set for this sample report.", GREY
T_OT, T_FPQ, T_TAT, T_ERR, T_FTE = "Targets!$B$4", "Targets!$B$5", "Targets!$B$6", "Targets!$B$7", "Targets!$B$8"

# ------------------------------------------------------------------ Jobs
js = wb.create_sheet("Jobs")
cols = ["JobID", "Client", "Team", "JobType", "ReceivedDate", "DueDate", "DeliveredDate", "Pages", "QAErrors", "Rework", "Status",
        "Received Month", "Delivered Month", "Turnaround Days", "On Time"]
header(js, 1, cols, [11, 22, 18, 20, 13, 13, 13, 8, 9, 8, 13, 11, 11, 11, 9])
for r, row in enumerate(jobs.itertuples(index=False), start=2):
    js.cell(row=r, column=1, value=row.JobID)
    js.cell(row=r, column=2, value=row.Client)
    js.cell(row=r, column=3, value=row.Team)
    js.cell(row=r, column=4, value=row.JobType)
    for c, v in ((5, row.ReceivedDate), (6, row.DueDate), (7, row.DeliveredDate)):
        if v:
            cell = js.cell(row=r, column=c, value=pd.Timestamp(v).to_pydatetime())
            cell.number_format = "yyyy-mm-dd"
    js.cell(row=r, column=8, value=int(row.Pages))
    js.cell(row=r, column=9, value=int(row.QAErrors))
    js.cell(row=r, column=10, value=int(row.Rework))
    js.cell(row=r, column=11, value=row.Status)
    js.cell(row=r, column=12, value=f'=TEXT(E{r},"yyyy-mm")')
    js.cell(row=r, column=13, value=f'=IF(G{r}="","",TEXT(G{r},"yyyy-mm"))')
    js.cell(row=r, column=14, value=f'=IF(G{r}="","",G{r}-E{r})')
    js.cell(row=r, column=15, value=f'=IF(G{r}="","",IF(G{r}<=F{r},1,0))')
js.freeze_panes = "A2"
js.auto_filter.ref = f"A1:O{NJ}"

# ------------------------------------------------------------------ Staffing / Escalations
st = wb.create_sheet("Staffing")
header(st, 1, ["Month", "Team", "FTE", "PlannedFTE"], [10, 20, 8, 12])
for r, row in enumerate(staff.itertuples(index=False), start=2):
    for c, v in enumerate(row, start=1):
        st.cell(row=r, column=c, value=v).font = B
es = wb.create_sheet("Escalations")
header(es, 1, ["EscalationID", "Date", "Client", "Team", "JobID", "Reason", "Month"], [12, 12, 22, 18, 11, 16, 10])
for r, row in enumerate(esc.itertuples(index=False), start=2):
    es.cell(row=r, column=1, value=row.EscalationID)
    es.cell(row=r, column=2, value=pd.Timestamp(row.Date).to_pydatetime()).number_format = "yyyy-mm-dd"
    for c, v in ((3, row.Client), (4, row.Team), (5, row.JobID), (6, row.Reason)):
        es.cell(row=r, column=c, value=v)
    es.cell(row=r, column=7, value=f'=TEXT(B{r},"yyyy-mm")')

# ------------------------------------------------------------------ Monthly KPIs
mk = wb.create_sheet("Monthly KPIs", 1)
mk["A1"], mk["A1"].font = "Monthly KPIs - all teams", H1
mk["A2"], mk["A2"].font = "Every value is a formula over the Jobs, Staffing and Escalations sheets. Delivery KPIs use the month a job was delivered.", GREY
hdr = ["Month", "Month end", "Jobs received", "Jobs delivered", "On-time %", "Avg turnaround (days)", "First-pass quality %",
       "Errors per 1,000 pages", "Pages delivered", "FTE", "Pages per FTE", "Escalations", "Open backlog at month end"]
header(mk, 4, hdr, [10, 12, 11, 11, 10, 13, 12, 13, 12, 8, 11, 11, 14])
for i, m in enumerate(months):
    r = 5 + i
    mk.cell(row=r, column=1, value=m)
    mk.cell(row=r, column=2, value=f"=EOMONTH(DATE(LEFT(A{r},4),RIGHT(A{r},2),1),0)").number_format = "yyyy-mm-dd"
    mk.cell(row=r, column=3, value=f"=COUNTIFS({J['recm']},A{r})")
    mk.cell(row=r, column=4, value=f"=COUNTIFS({J['delm']},A{r})")
    mk.cell(row=r, column=5, value=f'=IFERROR(SUMIFS({J["ontime"]},{J["delm"]},A{r})/D{r},"")').number_format = "0.0%"
    mk.cell(row=r, column=6, value=f'=IFERROR(AVERAGEIFS({J["tat"]},{J["delm"]},A{r}),"")').number_format = "0.0"
    mk.cell(row=r, column=7, value=f'=IFERROR(1-COUNTIFS({J["delm"]},A{r},{J["rework"]},1)/D{r},"")').number_format = "0.0%"
    mk.cell(row=r, column=8, value=f'=IFERROR(SUMIFS({J["err"]},{J["delm"]},A{r})/SUMIFS({J["pages"]},{J["delm"]},A{r})*1000,"")').number_format = "0.00"
    mk.cell(row=r, column=9, value=f"=SUMIFS({J['pages']},{J['delm']},A{r})").number_format = "#,##0"
    mk.cell(row=r, column=10, value=f"=SUMIFS(Staffing!$C$2:$C${NS},Staffing!$A$2:$A${NS},A{r})").number_format = "0.0"
    mk.cell(row=r, column=11, value=f'=IFERROR(I{r}/J{r},"")').number_format = "#,##0"
    mk.cell(row=r, column=12, value=f"=COUNTIFS(Escalations!$G$2:$G${NE},A{r})")
    mk.cell(row=r, column=13, value=f'=COUNTIFS({J["rec"]},"<="&B{r},{J["del"]},">"&B{r})'
                                    f'+COUNTIFS({J["rec"]},"<="&B{r},{J["status"]},"In progress")')
    for c in range(1, 14):
        mk.cell(row=r, column=c).font, mk.cell(row=r, column=c).border = B, BOX
last = 4 + len(months)
tr = last + 1
mk.cell(row=tr, column=1, value="Total / avg").font = BB
for c, f in ((3, f"=SUM(C5:C{last})"), (4, f"=SUM(D5:D{last})"),
             (5, f"=SUM({J['ontime']})/D{tr}"), (6, f"=AVERAGE({J['tat']})"),
             (7, f"=1-COUNTIFS({J['rework']},1,{J['status']},\"Delivered\")/D{tr}"),
             (8, f"=SUMIFS({J['err']},{J['status']},\"Delivered\")/SUMIFS({J['pages']},{J['status']},\"Delivered\")*1000"),
             (9, f"=SUM(I5:I{last})"), (12, f"=SUM(L5:L{last})")):
    cell = mk.cell(row=tr, column=c, value=f)
    cell.font, cell.fill, cell.border = BB, LIGHT, BOX
    cell.number_format = mk.cell(row=5, column=c).number_format
mk.cell(row=tr + 2, column=1, value="Note: jobs received in a month can be delivered the next month, so 'received' and 'delivered' differ. "
        "Jobs marked 'Needs review' (bad dates) are excluded from delivery KPIs.").font = GREY
rngE, rngG = f"E5:E{last}", f"G5:G{last}"
for rg, t in ((rngE, T_OT), (rngG, T_FPQ)):
    mk.conditional_formatting.add(rg, CellIsRule(operator="greaterThanOrEqual", formula=[t], fill=GOOD))
    mk.conditional_formatting.add(rg, CellIsRule(operator="between", formula=[f"{t}-0.05", f"{t}-0.0001"], fill=WARN))
    mk.conditional_formatting.add(rg, CellIsRule(operator="lessThan", formula=[f"{t}-0.05"], fill=BAD))
mk.conditional_formatting.add(f"F5:F{last}", CellIsRule(operator="greaterThan", formula=[T_TAT], fill=BAD))
mk.conditional_formatting.add(f"H5:H{last}", CellIsRule(operator="greaterThan", formula=[T_ERR], fill=BAD))
mk.freeze_panes = "B5"

ch = LineChart()
ch.title, ch.y_axis.title, ch.height, ch.width = "On-time % and first-pass quality %", "%", 8, 18
ch.add_data(Reference(mk, min_col=5, min_row=4, max_row=last), titles_from_data=True)
ch.add_data(Reference(mk, min_col=7, min_row=4, max_row=last), titles_from_data=True)
ch.set_categories(Reference(mk, min_col=1, min_row=5, max_row=last))
ch.y_axis.number_format, ch.y_axis.scaling.min, ch.y_axis.scaling.max = "0%", 0.5, 1.0
bc = BarChart()
bc.title, bc.height, bc.width = "Jobs received vs delivered", 8, 18
bc.add_data(Reference(mk, min_col=3, max_col=4, min_row=4, max_row=last), titles_from_data=True)
bc.set_categories(Reference(mk, min_col=1, min_row=5, max_row=last))
for c in (ch, bc):
    c.roundedCorners = False
    c.legend.position = "b"
    c.legend.overlay = False
    c.title.overlay = False
    c.x_axis.delete = False
    c.y_axis.delete = False
mk.add_chart(ch, f"A{tr + 4}")
mk.add_chart(bc, f"H{tr + 4}")

# ------------------------------------------------------------------ Team Scorecard
sc = wb.create_sheet("Team Scorecard", 2)
sc["A1"], sc["A1"].font = "Team scorecard", H1
sc["A3"], sc["A3"].font = "Select month:", BB
sc["C3"] = "2024-10"
sc["C3"].fill, sc["C3"].font = YEL, BLUE
dv = DataValidation(type="list", formula1=f"='Monthly KPIs'!$A$5:$A${last}", allow_blank=False)
sc.add_data_validation(dv)
dv.add("C3")
sc["D3"] = '=IFERROR(INDEX(\'Monthly KPIs\'!$A$5:$A$' + str(last) + ',MATCH(C3,\'Monthly KPIs\'!$A$5:$A$' + str(last) + ',0)-1),"")'
sc["E3"], sc["E3"].font = "<- previous month (for comparison)", GREY
h = ["Team", "Received", "Delivered", "On-time %", "On-time % prev month", "Change (pts)", "Status", "First-pass quality %",
     "Avg turnaround (days)", "Errors per 1,000 pages", "FTE", "Planned FTE", "Staffing vs plan %", "Pages per FTE", "Escalations"]
header(sc, 5, h, [20, 10, 10, 10, 12, 10, 12, 12, 12, 12, 8, 9, 11, 10, 11])
for i, t in enumerate(teams + ["All teams"]):
    r = 6 + i
    tcrit = f',{J["team"]},$A{r}' if t != "All teams" else ""
    scrit = f",Staffing!$B$2:$B${NS},$A{r}" if t != "All teams" else ""
    ecrit = f",Escalations!$D$2:$D${NE},$A{r}" if t != "All teams" else ""
    sc.cell(row=r, column=1, value=t)
    sc.cell(row=r, column=2, value=f"=COUNTIFS({J['recm']},$C$3{tcrit})")
    sc.cell(row=r, column=3, value=f"=COUNTIFS({J['delm']},$C$3{tcrit})")
    sc.cell(row=r, column=4, value=f'=IFERROR(SUMIFS({J["ontime"]},{J["delm"]},$C$3{tcrit})/C{r},"")').number_format = "0.0%"
    sc.cell(row=r, column=5, value=f'=IFERROR(SUMIFS({J["ontime"]},{J["delm"]},$D$3{tcrit})/COUNTIFS({J["delm"]},$D$3{tcrit}),"")').number_format = "0.0%"
    sc.cell(row=r, column=6, value=f'=IFERROR((D{r}-E{r})*100,"")').number_format = "+0.0;-0.0;0.0"
    sc.cell(row=r, column=7, value=f'=IF(D{r}="","",IF(D{r}>={T_OT},"On target",IF(D{r}>={T_OT}-0.05,"Watch","Off target")))')
    sc.cell(row=r, column=8, value=f'=IFERROR(1-COUNTIFS({J["delm"]},$C$3,{J["rework"]},1{tcrit})/C{r},"")').number_format = "0.0%"
    sc.cell(row=r, column=9, value=f'=IFERROR(AVERAGEIFS({J["tat"]},{J["delm"]},$C$3{tcrit}),"")').number_format = "0.0"
    sc.cell(row=r, column=10, value=f'=IFERROR(SUMIFS({J["err"]},{J["delm"]},$C$3{tcrit})/SUMIFS({J["pages"]},{J["delm"]},$C$3{tcrit})*1000,"")').number_format = "0.00"
    sc.cell(row=r, column=11, value=f"=SUMIFS(Staffing!$C$2:$C${NS},Staffing!$A$2:$A${NS},$C$3{scrit})").number_format = "0.0"
    sc.cell(row=r, column=12, value=f"=SUMIFS(Staffing!$D$2:$D${NS},Staffing!$A$2:$A${NS},$C$3{scrit})").number_format = "0.0"
    sc.cell(row=r, column=13, value=f'=IFERROR(K{r}/L{r},"")').number_format = "0.0%"
    sc.cell(row=r, column=14, value=f'=IFERROR(SUMIFS({J["pages"]},{J["delm"]},$C$3{tcrit})/K{r},"")').number_format = "#,##0"
    sc.cell(row=r, column=15, value=f"=COUNTIFS(Escalations!$G$2:$G${NE},$C$3{ecrit})")
    for c in range(1, 16):
        cell = sc.cell(row=r, column=c)
        cell.font, cell.border = (BB if t == "All teams" else B), BOX
        if t == "All teams":
            cell.fill = LIGHT
lr = 6 + len(teams)
sc.conditional_formatting.add(f"G6:G{lr}", FormulaRule(formula=['G6="On target"'], fill=GOOD))
sc.conditional_formatting.add(f"G6:G{lr}", FormulaRule(formula=['G6="Watch"'], fill=WARN))
sc.conditional_formatting.add(f"G6:G{lr}", FormulaRule(formula=['G6="Off target"'], fill=BAD))
sc.conditional_formatting.add(f"M6:M{lr}", CellIsRule(operator="lessThan", formula=[T_FTE], fill=BAD))
sc.conditional_formatting.add(f"H6:H{lr}", CellIsRule(operator="lessThan", formula=[T_FPQ], fill=BAD))
sc[f"A{lr + 2}"], sc[f"A{lr + 2}"].font = "Status: On target = on-time % at or above target; Watch = within 5 points; Off target = more than 5 points below.", GREY

# ------------------------------------------------------------------ DQ Checklist
dq = wb.create_sheet("DQ Checklist")
dq["A1"], dq["A1"].font = "Data-quality checklist", H1
dq["A2"], dq["A2"].font = "Part A runs automatically on the Jobs data in this workbook. Part B is the monthly manual checklist (fill in Done / By).", GREY
header(dq, 4, ["#", "Automated check (Part A)", "Rows failing", "Result"], [5, 70, 13, 10])
checks = [
    ("Duplicate JobID (data sorted by JobID)", f"=SUMPRODUCT(--(Jobs!$A$3:$A${NJ}=Jobs!$A$2:$A${NJ - 1}))"),
    ("Team is blank", f"=COUNTBLANK({J['team']})"),
    ("Team not in the staffing team list", f"=SUMPRODUCT(--(COUNTIF(Staffing!$B$2:$B${NS},{J['team']})=0))"),
    ("Pages zero or negative", f'=COUNTIF({J["pages"]},"<=0")'),
    ("Due date before received date", f"=SUMPRODUCT(--({J['due']}<{J['rec']}))"),
    ("Delivered date before received date", f'=SUMPRODUCT(({J["del"]}<>"")*({J["del"]}<{J["rec"]}))'),
    ("Status 'Delivered' but no delivered date", f'=COUNTIFS({J["status"]},"Delivered",{J["del"]},"")'),
    ("Delivered date present but status not 'Delivered'", f'=COUNTIFS({J["status"]},"<>Delivered",{J["del"]},"<>")'),
    ("Rework flag not 0 or 1", f'=COUNTIFS({J["rework"]},"<>0",{J["rework"]},"<>1")'),
    ("Staffing rows missing (expect 3 teams x months)", f"={len(months) * len(teams)}-COUNTA(Staffing!$A$2:$A${NS})"),
    ("Jobs marked 'Needs review' (excluded from delivery KPIs)", f'=COUNTIF({J["status"]},"Needs review")'),
]
for i, (name, f) in enumerate(checks, start=1):
    r = 4 + i
    dq.cell(row=r, column=1, value=i)
    dq.cell(row=r, column=2, value=name)
    dq.cell(row=r, column=3, value=f)
    dq.cell(row=r, column=4, value=f'=IF(C{r}=0,"Pass",IF(ISNUMBER(SEARCH("Needs review",B{r})),"Info","Fail"))')
    for c in range(1, 5):
        dq.cell(row=r, column=c).font, dq.cell(row=r, column=c).border = B, BOX
lr = 4 + len(checks)
dq.conditional_formatting.add(f"D5:D{lr}", FormulaRule(formula=['D5="Pass"'], fill=GOOD))
dq.conditional_formatting.add(f"D5:D{lr}", FormulaRule(formula=['D5="Fail"'], fill=BAD))
dq.conditional_formatting.add(f"D5:D{lr}", FormulaRule(formula=['D5="Info"'], fill=WARN))
log = pd.read_csv(os.path.join(HERE, "..", "01. Dataset", "Raw Dataset", "Jobs.csv"), nrows=0)  # touch path; counts below come from the log
r0 = lr + 2
dq.cell(row=r0, column=2, value="Issues found in the raw extract and fixed by clean_dataset.py (source: 01. Dataset/Cleaning Log.md)").font = BB
for i, line in enumerate([l for l in open(os.path.join(HERE, "..", "01. Dataset", "Cleaning Log.md"), encoding="utf-8")
                          if l.startswith("| Jobs")], start=1):
    _, t, a, n, _ = [x.strip() for x in line.split("|")]
    dq.cell(row=r0 + i, column=2, value=a).font = B
    dq.cell(row=r0 + i, column=3, value=int(n.replace(",", ""))).font = B
r1 = r0 + 9
header_row = r1
for i, c in enumerate(["#", "Manual monthly check (Part B)", "Owner", "Done (Y/N)", "By", "Notes"], start=1):
    cell = dq.cell(row=header_row, column=i, value=c)
    cell.font, cell.fill, cell.border = H2, HEAD, BOX
dq.column_dimensions["E"].width, dq.column_dimensions["F"].width = 12, 40
manual = [("Extract covers the full month (first and last job dates checked)", "MIS analyst"),
          ("Row count within +/-20% of last month; large changes explained", "MIS analyst"),
          ("Part A automated checks all Pass (or failures explained)", "MIS analyst"),
          ("New clients / job types mapped to the right team", "Team leads"),
          ("Staffing (FTE) confirmed with HR / resource planner", "Operations manager"),
          ("Escalations list matches the client-service log", "Client services"),
          ("KPI movements over 5 points explained in the commentary", "Operations manager"),
          ("Report reviewed by a second person before release", "Reviewer")]
for i, (c, o) in enumerate(manual, start=1):
    r = header_row + i
    vals = [i, c, o, "N" if i > 1 else "Y", "" if i > 1 else "Example: B. R.", "" if i > 1 else "Example row - replace each month"]
    for j, v in enumerate(vals, start=1):
        cell = dq.cell(row=r, column=j, value=v)
        cell.font, cell.border = (BLUE if j in (4, 5, 6) else B), BOX

# ------------------------------------------------------------------ KPI Definitions
kd = wb.create_sheet("KPI Definitions")
kd["A1"], kd["A1"].font = "KPI definitions", H1
header(kd, 3, ["KPI", "Definition", "Excel logic (Monthly KPIs)", "Owner"], [26, 55, 70, 18])
defs = [("Jobs received", "Jobs with a received date in the month", "COUNTIFS(Received Month = month)", "Operations"),
        ("Jobs delivered", "Jobs with a delivered date in the month", "COUNTIFS(Delivered Month = month)", "Operations"),
        ("On-time %", "Share of delivered jobs delivered on or before the due date", "SUMIFS(On Time) / Jobs delivered", "Operations"),
        ("Avg turnaround (days)", "Average calendar days from received to delivered", "AVERAGEIFS(Turnaround Days)", "Operations"),
        ("First-pass quality %", "Share of delivered jobs that needed no rework", "1 - COUNTIFS(Rework = 1) / Jobs delivered", "Quality"),
        ("Errors per 1,000 pages", "QA errors found per 1,000 pages delivered", "SUMIFS(QAErrors) / SUMIFS(Pages) x 1,000", "Quality"),
        ("Pages per FTE", "Pages delivered per full-time equivalent in the month", "Pages delivered / SUMIFS(Staffing FTE)", "Operations"),
        ("Staffing vs plan %", "Actual FTE / planned FTE", "SUMIFS(FTE) / SUMIFS(PlannedFTE)", "Resource planning"),
        ("Escalations", "Client escalations logged in the month", "COUNTIFS(Escalations Month = month)", "Client services"),
        ("Open backlog at month end", "Jobs received by month end and not delivered by month end",
         "COUNTIFS(Received <= month end, Delivered > month end) + COUNTIFS(Received <= month end, Status = In progress)", "Operations")]
for i, row in enumerate(defs, start=4):
    for j, v in enumerate(row, start=1):
        cell = kd.cell(row=i, column=j, value=v)
        cell.font, cell.border, cell.alignment = B, BOX, Alignment(wrap_text=True, vertical="top")

for w in wb.worksheets:
    w.sheet_view.showGridLines = w.title in ("Jobs", "Staffing", "Escalations")
wb.save(OUT)
print("saved", OUT, "jobs rows", len(jobs))
