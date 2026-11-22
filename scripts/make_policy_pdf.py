"""Build docs/Fee_Payment_Policy.pdf – the fictional policy used as Copilot Studio knowledge."""
from pathlib import Path

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

OUT = Path(__file__).resolve().parents[1] / "docs" / "Fee_Payment_Policy.pdf"
ss = getSampleStyleSheet()
body = ParagraphStyle("body", parent=ss["BodyText"], fontSize=10.5, leading=14.5, spaceAfter=6)
h1 = ParagraphStyle("h1", parent=ss["Title"], fontSize=18, spaceAfter=4)
h2 = ParagraphStyle("h2", parent=ss["Heading2"], fontSize=12.5, spaceBefore=10, spaceAfter=4,
                    textColor=colors.HexColor("#1F3A5F"))
note = ParagraphStyle("note", parent=body, fontSize=9, textColor=colors.HexColor("#52606D"))

story = [
    Paragraph("Student Fee Payment Policy", h1),
    Paragraph("Sample document for a portfolio project. Fictional – not the policy of any real "
              "institution. Version 1.0, effective 1 January 2025.", note),
    Spacer(1, 6),
    Paragraph("Section 1 – Fees and due dates", h2),
    Paragraph("The following fees are charged every semester. Due dates are counted from the first "
              "day of the semester.", body),
]
fee_table = Table([
    ["Fee", "Amount (LKR)", "Who pays", "Due"],
    ["Registration", "2,500", "All students", "On the semester start date"],
    ["Hostel", "18,000", "Hostel residents", "14 days after semester start"],
    ["Lab", "3,000", "All faculties except Business Studies", "60 days after semester start"],
    ["Library", "1,000", "All students", "90 days after semester start"],
    ["Exam", "2,000", "Applied Sciences, Technology, Medicine", "126 days after semester start"],
    ["Exam", "1,500", "Business Studies, Agriculture", "126 days after semester start"],
], colWidths=[24 * mm, 26 * mm, 66 * mm, 52 * mm])
fee_table.setStyle(TableStyle([
    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1F3A5F")),
    ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
    ("FONTSIZE", (0, 0), (-1, -1), 9.5),
    ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#C8D1DA")),
    ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F3F6F9")]),
]))
story += [fee_table, Spacer(1, 6)]

sections = [
    ("Section 2 – Payment methods",
     ["Fees may be paid online through the student portal, by bank deposit to the university "
      "collection account, or in cash at the Bursar's office.",
      "A cash payment of more than LKR 15,000 in a single transaction is accepted only with prior "
      "written approval from the Bursar. Cash payments above this limit that are recorded without "
      "approval are flagged for review."]),
    ("Section 3 – Receipts",
     ["Every payment must be recorded with an official receipt number on the day it is received.",
      "A payment recorded without a receipt number is reviewed by the finance office within 5 "
      "working days. The finance office confirms the payment with the student or the bank and "
      "issues the missing receipt."]),
    ("Section 4 – Duplicate payments",
     ["If the same fee is paid twice for the same amount, the second payment is treated as a "
      "duplicate. The duplicate amount is refunded to the student within 14 days after the Bursar "
      "has verified it."]),
    ("Section 5 – Overpayments",
     ["A single payment must not be larger than the fee it pays. Any amount above the fee is "
      "investigated as a possible recording error. Once confirmed, the excess is refunded or "
      "credited to the student's next fee, at the student's choice."]),
    ("Section 6 – Overdue fees",
     ["A fee is overdue when it is not paid in full by its due date. The finance office sends a "
      "reminder when a fee becomes overdue.",
      "If any fee is more than 90 days overdue, the student's exam admission card is withheld "
      "until the fee is paid in full."]),
    ("Section 7 – Recording standards",
     ["Payments must be recorded with the actual date the money was received. Entries dated in "
      "the future are not permitted and are corrected by the finance office.",
      "Each payment record shows the staff member who recorded it, or 'Online Portal' for online "
      "payments."]),
    ("Section 8 – Contact",
     ["Questions not covered by this policy should be sent to the finance office."]),
]
for title, paras in sections:
    story.append(Paragraph(title, h2))
    story += [Paragraph(p, body) for p in paras]

OUT.parent.mkdir(parents=True, exist_ok=True)
SimpleDocTemplate(str(OUT), pagesize=A4, leftMargin=20 * mm, rightMargin=20 * mm,
                  topMargin=18 * mm, bottomMargin=18 * mm,
                  title="Student Fee Payment Policy (sample)").build(story)
print("Wrote", OUT)
