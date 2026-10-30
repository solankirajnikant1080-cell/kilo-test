#!/usr/bin/env python3
"""
form8-template.xlsx બનાવે છે — CHC HMIS Form-8 માસિક આંકડા માટે તૈયાર Excel ટેમ્પલેટ.

    python3 form8/make_template.py            # -> form8/form8-template.xlsx

લક્ષણો:
  * ઉપર હેડર (ફેસિલિટી / તાલુકો / મહિનો / ભરનાર)
  * દરેક દિવસ માટે પંક્તિ (1..31), આંકડા માટે નંબર ફોર્મેટ
  * નીચે "કુલ" પંક્તિમાં SUM() ફોર્મ્યુલા (આપોઆપ ગણાય) + રોજનું સરેરાશ
  * Freeze pane + AutoFilter + A4 લેન્ડસ્કેપ પ્રિન્ટ ફિટ
"""
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
import datetime
import os

# (કોલમનું નામ, પહોળાઈ, પ્રકાર)  પ્રકાર: "sr" | "date" | "num" | "text"
COLUMNS = [
    ("ક્રમ",          6,  "sr"),
    ("તારીખ",        9,  "date"),
    ("OPD કુલ દર્દી", 14, "num"),
    ("OPD નવા",      12, "num"),
    ("OPD ફૉલો-અપ",  13, "num"),
    ("IPD દાખલ",     11, "num"),
    ("IPD કુલ ભરતી", 13, "num"),
    ("રેફર કરેલા",   11, "num"),
    ("લેબ ટેસ્ટ",    11, "num"),
    ("ડિલિવરી",      10, "num"),
    ("ટીકરણ (ડોઝ)",  13, "num"),
    ("નોંધ",         24, "text"),
]

THIN = Side(style="thin", color="9BB0C2")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD_FILL = PatternFill("solid", fgColor="0B6B57")
HEAD_FONT = Font(bold=True, color="FFFFFF", size=11)
TOTAL_FILL = PatternFill("solid", fgColor="E3F1EC")
LABEL_FONT = Font(bold=True, size=11)


def build(path, year=None, month=None):
    today = datetime.date.today()
    year = year or today.year
    month = month or today.month
    days = (datetime.date(year + month % 12, month % 12 + 1, 1) - datetime.timedelta(days=1)).day

    wb = Workbook()
    ws = wb.active
    ws.title = "Form-8 %d-%02d" % (year, month)

    ncols = len(COLUMNS)
    last_col = get_column_letter(ncols)

    # --- ટાઇટલ ---
    ws.merge_cells("A1:%s1" % last_col)
    ws["A1"] = "CHC HMIS — Form-8 માસિક આંકડા"
    ws["A1"].font = Font(bold=True, size=14, color="0B6B57")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    # --- હેડર ફીલ્ડ ---
    header_fields = [
        ("A2", "ફેસિલિટીનું નામ", "C2"),
        ("A3", "તાલુકો / બ્લોક", "C3"),
        ("A4", "મહિનો / વર્ષ", "C4"),
        ("E2", "ભરનારનું નામ", "G2"),
        ("E3", "પદ / હોદ્દો", "G3"),
        ("E4", "તારીખ", "G4"),
    ]
    for label_cell, label, value_cell in header_fields:
        ws[label_cell] = label
        ws[label_cell].font = LABEL_FONT
        ws[value_cell].border = BORDER
        ws[value_cell].alignment = Alignment(horizontal="left")
    ws["C4"] = "%02d/%d" % (month, year)
    ws["G4"] = today.strftime("%d/%m/%Y")
    # --- ટેબલ હેડર ---
    head_row = 6
    for i, (name, width, _kind) in enumerate(COLUMNS, start=1):
        cell = ws.cell(row=head_row, column=i, value=name)
        cell.fill = HEAD_FILL
        cell.font = HEAD_FONT
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = width
    ws.row_dimensions[head_row].height = 34

    # --- ડેટા પંક્તિઓ (દરેક દિવસ) ---
    first = head_row + 1
    last = first + days - 1
    for d in range(days):
        r = first + d
        ws.cell(row=r, column=1, value=d + 1)
        ws.cell(row=r, column=2, value=d + 1)
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER
            kind = COLUMNS[c - 1][2]
            if kind in ("sr", "date", "num"):
                cell.alignment = Alignment(horizontal="center")
                if kind == "num":
                    cell.number_format = "0"
            else:
                cell.alignment = Alignment(horizontal="left")
    # --- કુલ પંક્તિ (SUM ફોર્મ્યુલા) ---
    total_row = last + 1
    ws.cell(row=total_row, column=1, value="કુલ")
    ws.merge_cells(start_row=total_row, start_column=1, end_row=total_row, end_column=2)
    for c in range(1, ncols + 1):
        cell = ws.cell(row=total_row, column=c)
        cell.fill = TOTAL_FILL
        cell.font = Font(bold=True, size=11)
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center")
        if COLUMNS[c - 1][2] == "num":
            col = get_column_letter(c)
            cell.value = "=SUM(%s%d:%s%d)" % (col, first, col, last)
            cell.number_format = "0"

    # --- રોજનું સરેરાશ પંક્તિ ---
    avg_row = total_row + 1
    ws.cell(row=avg_row, column=1, value="રોજનું સરેરાશ")
    ws.merge_cells(start_row=avg_row, start_column=1, end_row=avg_row, end_column=2)
    for c in range(1, ncols + 1):
        cell = ws.cell(row=avg_row, column=c)
        cell.border = BORDER
        cell.alignment = Alignment(horizontal="center")
        cell.font = Font(italic=True, size=10, color="44586B")
        if COLUMNS[c - 1][2] == "num":
            col = get_column_letter(c)
            cell.value = "=IFERROR(ROUND(AVERAGE(%s%d:%s%d),1),0)" % (col, first, col, last)
            cell.number_format = "0.0"

    # --- વ્યૂ / પ્રિન્ટ સેટિંગ ---
    ws.freeze_panes = "A%d" % first
    ws.auto_filter.ref = "A%d:%s%d" % (head_row, last_col, last)
    ws.print_title_rows = "%d:%d" % (head_row, head_row)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True

    # --- નંબર કોલમમાં નેગેટિવ આંકડો ન ભરાય એ માટે વેલિડેશન ---
    dv = DataValidation(
        type="decimal", operator="greaterThanOrEqual", formula1="0",
        allow_blank=True, showErrorMessage=True,
        errorTitle="ખોટો આંકડો", error="ફક્ત 0 કે એથી મોટો આંકડો લખો."
    )
    dv.add("%s%d:%s%d" % ("C", first, get_column_letter(ncols), last))
    ws.add_data_validation(dv)

    os.makedirs(os.path.dirname(path) or ".", exist_ok=True)
    wb.save(path)
    return path


if __name__ == "__main__":
    out = os.path.join(os.path.dirname(os.path.abspath(__file__)), "form8-template.xlsx")
    print("saved:", build(out))

