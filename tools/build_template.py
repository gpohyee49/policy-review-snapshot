"""Builds Policy_Review_Template.xlsx — the input template + Excel-only dashboard.

Run:  python tools/build_template.py
The same file is what the browser app (Policy Review Snapshot.html) reads.
Sample client below is FICTIONAL, for format checking only.
"""
from datetime import date
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

OUT = Path(__file__).resolve().parent.parent / "Policy_Review_Template.xlsx"

# Blank copy for advisers to fill in (no sample rows); the app embeds this one for its download button
OUT_BLANK = Path(__file__).resolve().parent.parent / "Policy_Review_Template_blank.xlsx"

PURPLE = "6B3FA0"
PINK = "D94F8C"
LILAC = "F3ECFA"
CREAM = "FFFBF5"
WHITE = "FFFFFF"
thin = Side(style="thin", color="D9CDE8")
BOX = Border(left=thin, right=thin, top=thin, bottom=thin)

POLICY_COLS = [
    # (header, width, example-hint)
    ("Policy No", 16), ("Insurer", 16), ("Plan Name", 26), ("Policy Type", 18),
    ("Life Assured", 20), ("Policy Owner", 20), ("Status", 11),
    ("Commencement Date", 14), ("Expiry Date", 14),
    ("Premium (RM)", 13), ("Payment Mode", 13),
    ("Death (RM)", 13), ("TPD (RM)", 13), ("Critical Illness (RM)", 14),
    ("Early CI (RM)", 13), ("Medical Annual Limit (RM)", 15),
    ("Room & Board (RM/day)", 13), ("Personal Accident (RM)", 14),
    ("Waiver of Premium", 11), ("Nominee Name", 22), ("Nominee Relationship", 14),
    ("Nomination Type", 14), ("Purpose", 30), ("Notes", 28), ("Retirement Fund (RM)", 15),
    ("Trust Created", 12),
]
# Annual-premium helper column sits right after the input columns
from openpyxl.utils import get_column_letter
HELPER = get_column_letter(len(POLICY_COLS) + 1)
TYPES = "Whole Life,Term Life,Investment-Linked,Medical,Critical Illness,Personal Accident,Endowment / Savings,Other"
STATUSES = "Inforce,Lapsed,Paid-up,Surrendered,Matured"
MODES = "Monthly,Quarterly,Half-yearly,Yearly,Single"
NOMTYPES = "Beneficiary,Hibah,Trustee,Conditional Hibah,None"
# Malaysian life insurers and family takaful operators (BNM FSP directory / PIDM insurer members, checked Oct 2026)
INSURERS = ["AIA", "Allianz Life", "AmMetLife", "Etiqa Life", "FWD Insurance", "Generali Life", "Great Eastern Life",
            "Hong Leong Assurance", "Manulife", "MCIS", "Prudential Assurance", "Sun Life Malaysia", "Tokio Marine Life",
            "Zurich Life", "AIA PUBLIC Takaful", "AmMetLife Takaful", "Etiqa Family Takaful", "FWD Takaful",
            "Great Eastern Takaful", "Hong Leong MSIG Takaful", "Prudential BSN Takaful", "Sun Life Malaysia Takaful",
            "Takaful Malaysia (Keluarga)", "Takaful Ikhlas Family", "Zurich Takaful"]
PURPOSES = ("Family income protection,Medical & hospital bills,Critical illness recovery,Retirement,Children's education,"
            "Savings & wealth,Mortgage / loan protection,Estate & legacy,Accident cover,Business protection")

SAMPLE_CLIENT = [
    ("Client Name", "TAN MEI LING (SAMPLE)"),
    ("Age", 38),
    ("Annual Income (RM)", 96000),
    ("No. of Dependants", 2),
    ("Review Date", date(2026, 10, 29)),
    ("Prepared By", "Goh Poh Yee"),
]
GUIDELINES = [
    # label, value, note  (review standard: at least RM1m for life, CI, retirement and medical)
    ("Life cover target (RM)", 1000000, "Review standard"),
    ("TPD cover target (RM)", None, "Leave blank if no TPD standard"),
    ("CI cover target (RM)", 1000000, "Review standard"),
    ("Retirement target (RM)", 1000000, "Review standard"),
    ("Medical annual limit target (RM)", 1000000, "Review standard"),
    ("Room & board target (RM/day)", 200, "Adjust per client"),
    ("Premium budget (% of income)", 0.10, "Adjust per client"),
]

SAMPLE_POLICIES = [
    ["S-0001", "Great Eastern Life", "Whole Life Plan (sample)", "Whole Life", "TAN MEI LING", "TAN MEI LING",
     "Inforce", date(2015, 3, 15), date(2087, 3, 15), 250, "Monthly",
     200000, 200000, None, None, None, None, None, "Yes",
     "LEE WEI MING", "Husband", "Hibah", "Family income protection", ""],
    ["S-0002", "Great Eastern Life", "Medical Investment-Linked (sample)", "Medical", "TAN MEI LING", "TAN MEI LING",
     "Inforce", date(2018, 7, 1), date(2088, 7, 1), 280, "Monthly",
     50000, 50000, None, None, 150000, 150, None, "No",
     "", "", "", "Hospital & surgery bills", "Nominee not recorded - to confirm"],
    ["S-0003", "Other Insurer", "Critical Illness Plan (sample)", "Critical Illness", "TAN MEI LING", "TAN MEI LING",
     "Inforce", date(2020, 1, 10), date(2060, 1, 10), "RM 1,800.00", "Yearly",
     100000, None, 100000, None, None, None, None, "No",
     "WONG SIEW LAN", "Mother", "Beneficiary", "Income replacement if diagnosed with CI", ""],
    ["S-0004", "Other Insurer", "Personal Accident (sample)", "Personal Accident", "TAN MEI LING", "TAN MEI LING",
     "Inforce", date(2023, 5, 1), date(2027, 5, 1), 300, "Yearly",
     None, None, None, None, None, None, 100000, "No",
     "", "", "", "Accident cover", "Yearly renewable"],
    ["S-0005", "Great Eastern Life", "Education Endowment (sample)", "Endowment / Savings", "TAN MEI LING", "TAN MEI LING",
     "Lapsed", date(2012, 9, 1), date(2032, 9, 1), 100, "Monthly",
     30000, None, None, None, None, None, None, "No",
     "LEE JIA HUI", "Daughter", "Beneficiary", "Child education fund", "Lapsed 2021 - discuss reinstatement?"],
]

hdr_font = Font(bold=True, color=WHITE)
hdr_fill = PatternFill("solid", fgColor=PURPLE)
lilac_fill = PatternFill("solid", fgColor=LILAC)


def title(ws, text, sub):
    ws["A1"] = text
    ws["A1"].font = Font(bold=True, size=16, color=PURPLE)
    ws["A2"] = sub
    ws["A2"].font = Font(italic=True, size=10, color="7A6A8C")


def build_client(wb, blank=False):
    ws = wb.create_sheet("Client")
    title(ws, "Client Info  客户资料", "Fill column B. Review standards are YOUR planning benchmarks, not official figures.")
    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 40
    r = 4
    for label, val in SAMPLE_CLIENT:
        if blank:
            val = None
        ws.cell(r, 1, label).font = Font(bold=True)
        c = ws.cell(r, 2, val)
        c.fill = lilac_fill
        c.border = BOX
        if isinstance(val, date):
            c.number_format = "dd/mm/yyyy"
        elif label.endswith("(RM)"):
            c.number_format = "#,##0"
        r += 1
    r += 1
    ws.cell(r, 1, "Review standard  检视标准").font = Font(bold=True, size=12, color=PINK)
    r += 1
    for label, val, note in GUIDELINES:
        ws.cell(r, 1, label).font = Font(bold=True)
        c = ws.cell(r, 2, val)
        c.fill = lilac_fill
        c.border = BOX
        c.number_format = "0%" if "%" in label else "#,##0"
        ws.cell(r, 3, note).font = Font(italic=True, color="7A6A8C")
        r += 1
    # Retirement savings outside insurance (rows 21-22) - added to the retirement fund on the Dashboard
    r += 1
    ws.cell(r, 1, "Retirement savings  退休储蓄").font = Font(bold=True, size=12, color=PINK)
    r += 1
    for label, note in (("EPF / KWSP balance (RM)", "Latest statement / i-Akaun"), ("PRS balance (RM)", "Latest PRS statement")):
        ws.cell(r, 1, label).font = Font(bold=True)
        c = ws.cell(r, 2)
        c.fill = lilac_fill
        c.border = BOX
        c.number_format = "#,##0"
        ws.cell(r, 3, note).font = Font(italic=True, color="7A6A8C")
        r += 1
    # Contact (rows 25-26) and family members (header row 30, people from row 31) - for follow-up, not shown on the report
    r += 1
    ws.cell(r, 1, "Contact  联络资料").font = Font(bold=True, size=12, color=PINK)
    r += 1
    for label, fmt, note in (("Phone  电话", "@", "Not shown on the client report"), ("Date of birth  生日", "dd/mm/yyyy", "")):
        ws.cell(r, 1, label).font = Font(bold=True)
        c = ws.cell(r, 2)
        c.fill = lilac_fill
        c.border = BOX
        c.number_format = fmt
        ws.cell(r, 3, note).font = Font(italic=True, color="7A6A8C")
        r += 1
    r += 2
    ws.cell(r, 1, "Family members  家人").font = Font(bold=True, size=12, color=PINK)
    ws.cell(r, 3, "For birthdays and follow-up. Not shown on the client report.").font = Font(italic=True, color="7A6A8C")
    r += 1
    for j, h in enumerate(("Name  姓名", "Relationship  关系", "Birthday  生日", "Phone  电话", "Notes  备注"), 1):
        c = ws.cell(r, j, h)
        c.font = hdr_font
        c.fill = hdr_fill
    ws.column_dimensions["D"].width = 18
    ws.column_dimensions["E"].width = 30
    rel = DataValidation(type="list", formula1='"Spouse 配偶,Son 儿子,Daughter 女儿,Father 父亲,Mother 母亲,Brother 兄弟,Sister 姐妹,Other 其他"', allow_blank=True)
    rel.showErrorMessage = False
    ws.add_data_validation(rel)
    for rr in range(r + 1, r + 16):
        for j in range(1, 6):
            c = ws.cell(rr, j)
            c.border = BOX
            c.number_format = "dd/mm/yyyy" if j == 3 else "@" if j == 4 else "General"
        rel.add(f"B{rr}")
    assert r == 30, r
    return ws


def build_policies(wb, blank=False):
    ws = wb.create_sheet("Policies")
    for i, (h, w) in enumerate(POLICY_COLS, 1):
        c = ws.cell(1, i, h)
        c.font = hdr_font
        c.fill = hdr_fill
        c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        c.border = BOX
        ws.column_dimensions[c.column_letter].width = w
    # helper column: annualised premium (used by Dashboard formulas)
    helper = len(POLICY_COLS) + 1
    hc = ws.cell(1, helper, "Annual Premium (auto)")
    hc.font = hdr_font
    hc.fill = PatternFill("solid", fgColor=PINK)
    hc.alignment = Alignment(wrap_text=True, horizontal="center")
    ws.column_dimensions[hc.column_letter].width = 14
    ws.row_dimensions[1].height = 42
    ws.freeze_panes = "B2"

    for r, row in enumerate([] if blank else SAMPLE_POLICIES, 2):
        for c, v in enumerate(row, 1):
            cell = ws.cell(r, c, v)
            cell.border = BOX
            if isinstance(v, date):
                cell.number_format = "dd/mm/yyyy"
            elif isinstance(v, (int, float)):
                cell.number_format = "#,##0"
    date_cols = {"H", "I"}
    num_cols = {"J", "L", "M", "N", "O", "P", "Q", "R", "Y"}
    for r in range(2, 101):
        for c in range(1, len(POLICY_COLS) + 1):
            cell = ws.cell(r, c)
            cell.border = BOX
            if cell.column_letter in date_cols:
                cell.number_format = "dd/mm/yyyy"
            elif cell.column_letter in num_cols:
                cell.number_format = "#,##0"
        # VALUE() tolerates text like "1800"; text with "RM" shows as 0 until cleaned — the browser app cleans it for you.
        ws.cell(r, helper, f'=IF(J{r}="","",IFERROR(VALUE(SUBSTITUTE(SUBSTITUTE(J{r},"RM",""),",",""))*'
                           f'IFERROR(INDEX({{12,4,2,1,0}},MATCH(K{r},{{"Monthly","Quarterly","Half-yearly","Yearly","Single"}},0)),0),0))')
        ws.cell(r, helper).number_format = "#,##0"

    def dv(col, items):
        d = DataValidation(type="list", formula1=f'"{items}"', allow_blank=True)
        ws.add_data_validation(d)
        d.add(f"{col}2:{col}100")
    dv("D", TYPES)
    dv("G", STATUSES)
    dv("K", MODES)
    dv("S", "Yes,No")
    dv("V", NOMTYPES)
    dv("Z", "Yes,No,Not sure")
    # Purpose: suggested list, but any other text is allowed (no error popup)
    pur = DataValidation(type="list", formula1=f'"{PURPOSES}"', allow_blank=True, showErrorMessage=False)
    ws.add_data_validation(pur)
    pur.add("W2:W100")
    # Insurer: list kept on the hidden Lists sheet (too long for an inline list); other names allowed
    ins = DataValidation(type="list", formula1=f"=Lists!$A$2:$A${len(INSURERS)+1}", allow_blank=True, showErrorMessage=False)
    ws.add_data_validation(ins)
    ins.add("B2:B100")
    # highlight inforce rows with no nominee
    ws.conditional_formatting.add(
        "T2:T100", FormulaRule(formula=['AND($G2="Inforce",$T2="",$A2<>"")'],
                               fill=PatternFill("solid", fgColor="FDE2EC")))
    return ws


def build_dashboard(wb):
    ws = wb.create_sheet("Dashboard", 0)
    title(ws, "Policy Review Snapshot  保单检视摘要", "Auto-calculated from 'Client' and 'Policies'. Inforce and paid-up policies count toward coverage.")
    for col, w in zip("ABCDEF", (34, 18, 18, 18, 16, 4)):
        ws.column_dimensions[col].width = w
    ws["A3"] = "Client 客户:"
    ws["B3"] = "=Client!B4"
    ws["D3"] = "Review date 检视日期:"
    ws["E3"] = "=Client!B8"
    ws["E3"].number_format = "dd/mm/yyyy"
    for c in ("A3", "D3"):
        ws[c].font = Font(bold=True)

    heads = ["Coverage area  保障类别", "Existing 现有 (RM)", "Standard 标准 (RM)", "Gap 差距 (RM)", "Status 状态"]
    for i, h in enumerate(heads, 1):
        c = ws.cell(5, i, h)
        c.font = hdr_font
        c.fill = hdr_fill
        c.border = BOX
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    inf = 'Policies!$G$2:$G$100,"Inforce"'
    pu = 'Policies!$G$2:$G$100,"Paid-up"'
    SUM = lambda col: f"SUMIFS(Policies!${col}$2:${col}$100,{inf})+SUMIFS(Policies!${col}$2:${col}$100,{pu})"
    MAX = lambda col: f"MAX(_xlfn.MAXIFS(Policies!${col}$2:${col}$100,{inf}),_xlfn.MAXIFS(Policies!${col}$2:${col}$100,{pu}))"
    rows = [
        ("Life / Death  身故", f"={SUM('L')}", "=Client!B12"),
        ("TPD  完全永久残疾", f"={SUM('M')}", "=Client!B13"),
        ("Critical Illness  严重疾病", f"={SUM('N')}", "=Client!B14"),
        ("Retirement fund (plans + EPF + PRS)  退休金", f"={SUM('Y')}+N(Client!B21)+N(Client!B22)", "=Client!B15"),
        ("Medical annual limit (highest)  医药年限", f"={MAX('P')}", "=Client!B16"),
        ("Room & board (RM/day)  房租", f"={MAX('Q')}", "=Client!B17"),
        ("Personal Accident  意外", f"={SUM('R')}", None),
    ]
    for r, (label, existing, target) in enumerate(rows, 6):
        ws.cell(r, 1, label).font = Font(bold=True)
        ws.cell(r, 2, existing)
        if target:
            ws.cell(r, 3, target)
            ws.cell(r, 4, f"=MAX(0,C{r}-B{r})")
            ws.cell(r, 5, f'=IF(C{r}=0,"-",IF(B{r}>=C{r},"✔ On track 已达标","▲ Review 可加强"))')
        else:
            ws.cell(r, 3, "-")
            ws.cell(r, 4, "-")
            ws.cell(r, 5, "No standard set")
        for c in range(1, 6):
            ws.cell(r, c).border = BOX
            ws.cell(r, c).number_format = "#,##0"
        if r % 2 == 0:
            for c in range(1, 6):
                ws.cell(r, c).fill = lilac_fill

    ws["A14"] = "Premium  保费"
    ws["A14"].font = Font(bold=True, size=12, color=PINK)
    prem = [
        ("Total annual premium (inforce + paid-up)  年缴总保费", f"={SUM(HELPER)}"),
        ("Monthly equivalent  每月相当于", "=B15/12"),
        ("% of annual income  占年收入", "=IF(Client!B6>0,B15/Client!B6,0)"),
        ("Premium budget  保费预算", "=Client!B18"),
    ]
    for r, (label, f) in enumerate(prem, 15):
        ws.cell(r, 1, label).font = Font(bold=True)
        c = ws.cell(r, 2, f)
        c.number_format = "0.0%" if "%" in label or "budget" in label else "#,##0"
        c.border = BOX

    ws["A20"] = "Housekeeping  资料检查"
    ws["A20"].font = Font(bold=True, size=12, color=PINK)
    hk = [
        ("Inforce / paid-up policies  有效保单", f"=COUNTIFS({inf})+COUNTIFS({pu})"),
        ("Not inforce (excluded)  非有效保单", '=COUNTA(Policies!$A$2:$A$100)-B21'),
        ("Inforce without nominee  未提名", f'=COUNTIFS({inf},Policies!$T$2:$T$100,"",Policies!$A$2:$A$100,"<>")+COUNTIFS({pu},Policies!$T$2:$T$100,"",Policies!$A$2:$A$100,"<>")'),
    ]
    for r, (label, f) in enumerate(hk, 21):
        ws.cell(r, 1, label).font = Font(bold=True)
        ws.cell(r, 2, f).border = BOX

    chart = BarChart()
    chart.type = "bar"
    chart.title = "Existing vs review standard (RM)"
    chart.style = 10
    chart.y_axis.numFmt = "#,##0"
    chart.x_axis.delete = False          # show category labels (hidden by default)
    chart.y_axis.delete = False          # show the RM scale
    chart.legend.position = "r"
    chart.add_data(Reference(ws, min_col=2, min_row=5, max_col=3, max_row=9), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=6, max_row=9))
    chart.series[0].graphicalProperties.solidFill = PURPLE
    chart.series[1].graphicalProperties.solidFill = "E9B8D0"
    chart.height, chart.width = 9, 18
    ws.add_chart(chart, "A25")

    ws["A44"] = ("Disclaimer: Summary based on policy information provided as at review date. Benefits are subject to the "
                 "terms of each policy contract. Review standards are planning benchmarks, not recommendations or official figures.")
    ws["A44"].font = Font(italic=True, size=9, color="7A6A8C")
    ws["A44"].alignment = Alignment(wrap_text=True)
    ws.merge_cells("A44:E46")


def build_notes(wb):
    ws = wb.create_sheet("Review notes")
    title(ws, "Review notes  检视记录", "Filled by the app: areas to discuss, data checks and the change log.")
    for col, w in zip("ABCDE", (22, 60, 18, 26, 30)):
        ws.column_dimensions[col].width = w
    for r in range(4, 201):
        for c in range(1, 6):
            cell = ws.cell(r, c)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
    ws.sheet_properties.tabColor = "8A63C2"


def build_guide(wb, blank=False):
    ws = wb.create_sheet("Guide")
    ws.column_dimensions["A"].width = 110
    lines = [
        "HOW TO USE  使用方法",
        "1. Client sheet: fill client name, income, review date. Review standard is RM1m for Life, CI, Retirement, Medical - adjust if needed.",
        "2. Policies sheet: one row per policy. Use the dropdowns for Type, Status, Payment Mode.",
        "3. Leave a cell BLANK if you don't know it. Never type a guess - the app flags blanks for you to confirm.",
        "4. Dashboard sheet updates automatically (Excel-only version).",
        "5. For the visual PDF report: open 'Policy Review Snapshot.html', load this file, click 'Generate'.",
        "",
        "Pink nominee cell = inforce policy without nominee recorded.",
        "Only Inforce policies count toward coverage totals. Lapsed/Paid-up are listed separately.",
        "The rows marked (sample) are fictional - delete them before using with a real client.",
    ]
    if blank:
        lines[-1] = "Insurer and Purpose columns have dropdown lists - you can also type any other name."
    for r, t in enumerate(lines, 1):
        ws.cell(r, 1, t).font = Font(bold=(r == 1), size=14 if r == 1 else 11, color=PURPLE if r == 1 else "333333")


def build(out, blank=False):
    wb = Workbook()
    wb.remove(wb.active)
    build_dashboard(wb)
    build_client(wb, blank)
    build_policies(wb, blank)
    build_notes(wb)
    build_guide(wb, blank)
    wb.calculation.fullCalcOnLoad = True      # Excel recalculates the Dashboard when the file opens
    lists = wb.create_sheet("Lists")
    lists["A1"] = "Insurer (Malaysia life & family takaful)"
    for r, name in enumerate(INSURERS, 2):
        lists.cell(r, 1, name)
    lists.sheet_state = "hidden"
    wb["Dashboard"].sheet_properties.tabColor = PINK
    wb["Policies"].sheet_properties.tabColor = PURPLE
    wb.save(out)
    print(f"Saved {out}")


def main():
    build(OUT)
    build(OUT_BLANK, blank=True)


if __name__ == "__main__":
    main()
