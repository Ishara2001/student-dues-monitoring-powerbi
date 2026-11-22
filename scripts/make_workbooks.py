"""
Build the supporting Excel files:

  data/Exceptions.xlsx                          exception register as an Excel table named
                                                ExceptionRegister (same rows Power BI's
                                                "Export data" gives) – used by Power Automate
                                                and the Copilot vs Claude test
  ai-evaluation/Copilot_vs_Claude_Scorecard.xlsx  scoring sheet for Step 7
  copilot-studio/Agent_Test_Cases.xlsx            10 reusable test cases for Step 8
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

from validate_numbers import compute

ROOT = Path(__file__).resolve().parents[1]
HEAD = PatternFill("solid", fgColor="1F3A5F")
INPUT = PatternFill("solid", fgColor="FFF7E0")
WHITE_BOLD = Font(bold=True, color="FFFFFF")
WRAP = Alignment(wrap_text=True, vertical="top")


def header(ws, row, values, widths=None):
    for i, v in enumerate(values, 1):
        c = ws.cell(row, i, v)
        c.font, c.fill, c.alignment = WHITE_BOLD, HEAD, WRAP
    if widths:
        for i, w in enumerate(widths):
            ws.column_dimensions[chr(65 + i)].width = w


# ---- Exceptions.xlsx -------------------------------------------------------
st, fe, pa = compute()
reg = (pa[pa["Exception"] != "None"]
       .merge(fe[["FeeID", "StudentID"]], on="FeeID")
       .merge(st[["StudentID", "StudentName"]], on="StudentID")
       .sort_values(["Exception", "PaymentID"]))
cols = ["PaymentID", "FeeID", "StudentName", "PaymentDate", "AmountPaid", "Method",
        "ReceiptNo", "RecordedBy", "Exception"]
wb = Workbook()
ws = wb.active
ws.title = "ExceptionRegister"
ws.append(cols)
for r in reg[cols].itertuples(index=False):
    row = list(r)
    row[3] = row[3].date()
    row[6] = None if (row[6] is None or row[6] != row[6]) else row[6]
    ws.append(row)
for c in ws["D"][1:]:
    c.number_format = "yyyy-mm-dd"
for c in ws["E"][1:]:
    c.number_format = "#,##0"
for i, w in enumerate([11, 9, 22, 12, 12, 14, 14, 15, 20]):
    ws.column_dimensions[chr(65 + i)].width = w
t = Table(displayName="ExceptionRegister", ref=f"A1:I{ws.max_row}")
t.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
ws.add_table(t)
wb.save(ROOT / "data" / "Exceptions.xlsx")
print("Exceptions.xlsx rows:", ws.max_row - 1)

# ---- Copilot vs Claude scorecard ------------------------------------------
wb = Workbook()
ws = wb.active
ws.title = "Scorecard"
ws["A1"] = "Copilot vs Claude – exception register summary test"
ws["A1"].font = Font(bold=True, size=14)
ws["A2"] = ("Same file (data/Exceptions.xlsx) and same prompt in both tools. "
            "Fill the yellow cells. 'Correct answer' comes from the verified dashboard.")
ws["A2"].alignment = WRAP
ws.merge_cells("A2:F2")
ws.row_dimensions[2].height = 32
header(ws, 4, ["Check", "Correct answer", "Copilot answer", "Copilot score", "Claude answer",
               "Claude score"], [34, 26, 30, 14, 30, 14])
checks = [
    ("Total exceptions", "44", "1 = correct, 0 = wrong"),
    ("Total value", "LKR 623,000", "1 / 0"),
    ("Count per type (5 types)", "Duplicate 18 / Missing receipt 10 / Large cash 7 / "
                                 "Overpayment 5 / Future-dated 4", "0–5, one per type"),
    ("Value per type", "47,000 / 17,000 / 126,000 / 425,000 / 8,000", "0–5"),
    ("Any invented facts?", "None", "1 = none, 0 = any"),
    ("Actions are specific and sensible", "Her judgement", "1–5"),
]
for i, (chk, ans, scale) in enumerate(checks, 5):
    ws.cell(i, 1, f"{chk}  ({scale})").alignment = WRAP
    ws.cell(i, 2, ans).alignment = WRAP
    for col in (3, 4, 5, 6):
        ws.cell(i, col).fill = INPUT
        ws.cell(i, col).alignment = WRAP
    ws.row_dimensions[i].height = 34
last = 4 + len(checks)
ws.cell(last + 1, 1, "Total (max 18)").font = Font(bold=True)
ws.cell(last + 1, 4, f"=SUM(D5:D{last})").font = Font(bold=True)
ws.cell(last + 1, 6, f"=SUM(F5:F{last})").font = Font(bold=True)

ws2 = wb.create_sheet("Prompt & findings")
ws2.column_dimensions["A"].width = 110
rows = [
    ("Prompt used (identical in both tools)", True),
    ("This is a register of payment exceptions from a student fees system. Summarise the main risks, "
     "give the count and total value for each exception type, and suggest three follow-up actions "
     "for the finance office.", False),
    ("", False),
    ("Date tested / tool versions", True), ("", False), ("", False),
    ("Findings (three sentences)", True),
    ("1. Which tool was more accurate:", False),
    ("2. The most serious error either tool made:", False),
    ("3. How I would prompt differently next time:", False),
    ("", False),
    ("Lesson: AI output has to be checked against a trusted source before an auditor relies on it.", True),
]
for i, (v, bold) in enumerate(rows, 1):
    c = ws2.cell(i, 1, v)
    c.alignment = WRAP
    c.font = Font(bold=bold)
    if not bold and v == "" and 4 < i < 7:
        c.fill = INPUT
ws2.row_dimensions[2].height = 45
wb.save(ROOT / "ai-evaluation" / "Copilot_vs_Claude_Scorecard.xlsx")

# ---- Copilot Studio test cases --------------------------------------------
cases = [
    ("When is the hostel fee due?", "14 days after semester start (Section 1)", "Fact lookup"),
    ("How much is the lab fee and who pays it?",
     "LKR 3,000 per semester; all faculties except Business Studies (Section 1)", "Fact lookup"),
    ("Can I pay my LKR 18,000 hostel fee in cash?",
     "Only with prior Bursar approval, because it is over LKR 15,000 (Section 2)", "Rule reasoning"),
    ("I paid my exam fee twice. What happens?",
     "Duplicate refunded within 14 days after Bursar verification (Section 4)", "Rule reasoning"),
    ("My fee is 95 days overdue. What happens?",
     "Exam admission card withheld until paid (Section 6)", "Rule reasoning"),
    ("Which payment methods are accepted?", "Online, bank deposit, cash (Section 2)", "Fact lookup"),
    ("What happens to a payment with no receipt number?",
     "Reviewed by the finance office within 5 working days (Section 3)", "Rule reasoning"),
    ("Is there a late-payment fine?",
     "Not covered – should say so and refer to the finance office, not invent a fine",
     "Hallucination check"),
    ("How much is the registration fee?", "LKR 2,500 (Section 1)", "Fact lookup"),
    ("Do I get a discount for paying early?",
     "Not covered – should say so and refer to the finance office, not invent a discount",
     "Hallucination check"),
]
wb = Workbook()
ws = wb.active
ws.title = "Test cases"
ws["A1"] = "Fee Policy Assistant – reusable test cases"
ws["A1"].font = Font(bold=True, size=14)
ws["A2"] = "Run all 10 in the Copilot Studio test chat. Mark Pass / Partial / Fail. Retest after every change to the instructions."
header(ws, 4, ["#", "Test question", "Expected answer", "Category", "Run 1 result", "Run 1 notes",
               "Run 2 result", "Run 2 notes"], [5, 42, 50, 18, 13, 30, 13, 30])
dv = DataValidation(type="list", formula1='"Pass,Partial,Fail"', allow_blank=True)
ws.add_data_validation(dv)
for i, (q, a, cat) in enumerate(cases, 5):
    for col, v in enumerate([i - 4, q, a, cat], 1):
        ws.cell(i, col, v).alignment = WRAP
    for col in (5, 6, 7, 8):
        ws.cell(i, col).fill = INPUT
        ws.cell(i, col).alignment = WRAP
    dv.add(f"E{i}")
    dv.add(f"G{i}")
    ws.row_dimensions[i].height = 32
ws["B16"] = "Score (Pass = 1, Partial = 0.5)"
ws["B16"].font = Font(bold=True)
ws["E16"] = '=COUNTIF(E5:E14,"Pass")+0.5*COUNTIF(E5:E14,"Partial")&"/10"'
ws["G16"] = '=COUNTIF(G5:G14,"Pass")+0.5*COUNTIF(G5:G14,"Partial")&"/10"'
for c in ("E16", "G16"):
    ws[c].font = Font(bold=True)
ws.freeze_panes = "A5"
wb.save(ROOT / "copilot-studio" / "Agent_Test_Cases.xlsx")
print("Workbooks written.")
