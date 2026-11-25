"""
Unit-wise POS report (cash reconciliation) - template + VetBuddy auto-fill.

One workbook per month, one sheet per unit, one row per day:
  DATE | VETBUDDY | GOFRUGAL | TOTAL GROSS SALES | ADVANCE RECEIVED | CREDIT ADJUSTED |
  TOTAL RECEIVABLES | PAYTM | CASH TO BE DEPOSITED | ACTUAL DEPOSIT | VARIANCE | REMARK |
  CREDIT SALES | CREDIT CLOSER | ADVANCE ADJUSTMENT | PRIOR VARIANCE B/F

The automation only writes the VETBUDDY column (Final Sales - Return Total from
VetBuddy "Detailed Goods & Services usage by Patients"). Everything the team types
(Gofrugal, Paytm, deposits, credit, remarks) is left untouched.
"""
import calendar, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.utils import get_column_letter as L

GREEN = PatternFill("solid", fgColor="274E13")
PINK = PatternFill("solid", fgColor="EAD1F0")
YEL = PatternFill("solid", fgColor="FFF2CC")
AUTO = PatternFill("solid", fgColor="E2EFDA")
thin = Side(style="thin", color="808080")
B = Border(left=thin, right=thin, top=thin, bottom=thin)
WH = Font(name="Arial", bold=True, color="FFFFFF", size=10)
NUM = '#,##,##0;-#,##,##0;"-"'
HEADS = ["DATE", "VETBUDDY", "GOFRUGAL", "TOTAL GROSS SALES", "ADVANCE RECEIVED", "CREDIT ADJUSTED",
         "TOTAL RECEIVABLES", "PAYTM", "CASH TO BE DEPOSITED", "ACTUAL DEPOSIT", "VARIANCE", "REMARK",
         "CREDIT SALES", "CREDIT CLOSER", "ADVANCE ADJUSTMENT", "PRIOR VARIANCE B/F"]
INPUTS = {5, 8, 10, 12, 13, 14, 15, 16}
WIDTHS = [12, 12, 12, 12, 11, 11, 12, 11, 12, 11, 11, 28, 10, 10, 12, 11]
FIRST_ROW = 3


def file_name(month):
    return f"POS_Report_Unitwise_{month:%b_%Y}.xlsx"


def add_unit_sheets(wb, month, units, sheet_name=lambda u: u[:31], index=None, days=None):
    """Adds one template sheet per unit. Returns {unit: sheet title}."""
    days = days or [month.replace(day=d) for d in range(1, calendar.monthrange(month.year, month.month)[1] + 1)]
    titles = {}
    for k, u in enumerate(units):
        ws = wb.create_sheet(sheet_name(u), None if index is None else index + k)
        titles[u] = ws.title
        _unit_sheet(ws, u, days)
    return titles


def _unit_sheet(ws, u, days):
    ws.merge_cells("A1:A2"); ws.merge_cells("B1:P1")
    ws["A1"] = "DATE"; ws["B1"] = u.upper()
    for c in range(1, 17):
        for r in (1, 2):
            x = ws.cell(r, c); x.fill, x.border, x.font = GREEN, B, WH
            x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if c > 1:
            ws.cell(2, c, HEADS[c - 1])
        ws.column_dimensions[L(c)].width = WIDTHS[c - 1]
    ws["B1"].font = Font(name="Arial", bold=True, color="FFFFFF", size=26)
    ws.row_dimensions[1].height = 40; ws.row_dimensions[2].height = 58
    ws["B2"].comment = Comment("Filled automatically every morning from VetBuddy: "
                               "Net Final Sales = Final Sales - Return Total.", "Finance Automation")
    ws["P2"].comment = Comment("Carried-over variance from earlier periods; explain in Remark.",
                               "Finance Automation")
    for i, d in enumerate(days):
        r = FIRST_ROW + i
        ws.cell(r, 1, d).number_format = "DD-MM-YYYY"
        ws.cell(r, 4, f"=B{r}+C{r}")
        ws.cell(r, 6, f"=-(M{r}-N{r}+O{r})")
        ws.cell(r, 7, f"=D{r}+E{r}+F{r}")
        ws.cell(r, 9, f"=G{r}-H{r}")
        ws.cell(r, 11, f'=IF(COUNT(C{r},H{r},J{r})=0,"",I{r}-J{r}+P{r})')
        for c in range(1, 17):
            x = ws.cell(r, c); x.border = B
            x.font = Font(name="Arial", size=10, bold=c in (4, 6, 7, 9, 11),
                          color="0000FF" if c in INPUTS and c != 12 else "000000")
            if c > 1 and c != 12:
                x.number_format = NUM
            x.fill = PINK if c == 11 else YEL if c in INPUTS else AUTO if c in (2, 3) else PatternFill()
    t = FIRST_ROW + len(days)
    ws.cell(t, 1, "TOTAL")
    for c in range(2, 17):
        if c != 12:
            ws.cell(t, c, f"=SUM({L(c)}{FIRST_ROW}:{L(c)}{t - 1})").number_format = NUM
    for c in range(1, 17):
        ws.cell(t, c).fill, ws.cell(t, c).font, ws.cell(t, c).border = GREEN, WH, B
    ws.freeze_panes = "B3"
    notes = ["HOW TO USE",
             "Green columns (VETBUDDY, GOFRUGAL) fill automatically each morning. Yellow cells = team inputs. White/pink = formulas.",
             "Total Gross Sales = VetBuddy + Gofrugal",
             "Credit Adjusted = -(Credit Sales - Credit Closer + Advance Adjustment)",
             "Total Receivables = Total Gross Sales + Advance Received + Credit Adjusted",
             "Cash to be Deposited = Total Receivables - Paytm",
             "Variance = Cash to be Deposited - Actual Deposit + Prior Variance B/F"]
    for i, s in enumerate(notes):
        ws.cell(t + 2 + i, 1, s).font = Font(name="Arial", size=9, bold=(i == 0), italic=(i > 0))



def build_template(month, units):
    days = [month.replace(day=d) for d in range(1, calendar.monthrange(month.year, month.month)[1] + 1)]
    wb = Workbook()
    wb.remove(wb.active)
    add_unit_sheets(wb, month, units, days=days)
    s = wb.create_sheet("Summary", 0)
    s.merge_cells("A1:K1")
    s["A1"] = f"POS REPORT - UNIT-WISE SUMMARY ({month:%b %Y})".upper()
    s["A1"].font = Font(name="Arial", bold=True, size=16, color="FFFFFF"); s["A1"].fill = GREEN
    s["A1"].alignment = Alignment(horizontal="center", vertical="center"); s.row_dimensions[1].height = 32
    sh = ["UNIT"] + HEADS[1:11]
    for c, h in enumerate(sh, 1):
        x = s.cell(2, c, h); x.fill, x.font, x.border = GREEN, WH, B
        x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        s.column_dimensions[L(c)].width = 16 if c == 1 else 13
    s.row_dimensions[2].height = 45
    tot = FIRST_ROW + len(days)
    for i, u in enumerate(units):
        r = 3 + i
        s.cell(r, 1, u).font = Font(name="Arial", size=10, bold=True)
        for c in range(2, 12):
            x = s.cell(r, c, f"='{u[:31]}'!{L(c)}{tot}"); x.number_format = NUM
            x.font = Font(name="Arial", size=10, color="008000")
        for c in range(1, 12):
            s.cell(r, c).border = B
        s.cell(r, 11).fill = PINK
    r = 3 + len(units); s.cell(r, 1, "NETWORK TOTAL")
    for c in range(2, 12):
        s.cell(r, c, f"=SUM({L(c)}3:{L(c)}{r - 1})").number_format = NUM
    for c in range(1, 12):
        s.cell(r, c).fill, s.cell(r, c).font, s.cell(r, c).border = GREEN, WH, B
    s.freeze_panes = "B3"
    return wb


def fill_vetbuddy(wb, daywise, titles=None, col=2):
    """daywise: {(date, unit): amount}. Writes only one column (B = VETBUDDY, C = GOFRUGAL) of each unit sheet.
    titles: {unit: sheet title}; default = sheet title is the unit name."""
    written = 0
    by_title = {v: k for k, v in (titles or {}).items()}
    for ws in wb.worksheets:
        if ws.title == "Summary" or (titles and ws.title not in by_title):
            continue
        unit = by_title.get(ws.title, ws.title)
        for r in range(FIRST_ROW, ws.max_row + 1):
            v = ws.cell(r, 1).value
            if isinstance(v, dt.datetime):
                v = v.date()
            if not isinstance(v, dt.date):
                continue
            key = (v, unit)
            if key in daywise:
                ws.cell(r, col).value = round(daywise[key], 2)
                written += 1
    return written
