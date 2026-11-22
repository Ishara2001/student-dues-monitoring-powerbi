"""
Generate the fictional Student_Dues_Data.xlsx used by the Power BI dashboard.

The data is a snapshot "as at" 31 Oct 2026, amounts in LKR. It contains
deliberately planted data-quality problems (messy text) and policy breaches
(duplicates, overpayments, missing receipts, large cash payments and
future-dated entries) so the dashboard has something real to find.

The generator is deterministic (fixed seed) and is built so the published
check figures in the README come out exactly. Run:

    python scripts/generate_dataset.py
    python scripts/validate_numbers.py
"""
from __future__ import annotations

import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SEED = 20261031
REPORT_DATE = date(2026, 10, 31)
OUT = Path(__file__).resolve().parents[1] / "data" / "Student_Dues_Data.xlsx"

# ---- Targets (must match the guide / README check table) -------------------
N_STUDENTS = 240
N_FEES = 3550
N_PAYMENTS = 3137
TOTAL_BILLED = 12_359_000
OVERDUE_FEES = 364
OVERDUE_AMOUNT = 1_101_600
STUDENTS_OVERDUE = 189
NOT_YET_DUE_UNPAID = {1500: 72, 2000: 48}          # 204,000 outstanding, not yet due
OVERDUE_UNPAID, OVERDUE_PARTIAL = 330, 34

# ---- Reference data --------------------------------------------------------
FACULTIES = ["Applied Sciences", "Business Studies", "Agriculture", "Technology", "Medicine"]
FACULTY_WEIGHTS = [60, 50, 45, 45, 40]
# 12 messy spellings that clean down to the 5 faculties with Trim + Capitalize Each Word
FACULTY_VARIANTS = {
    "Applied Sciences": ["Applied Sciences", "applied sciences", "APPLIED SCIENCES "],
    "Business Studies": ["Business Studies", " business studies", "BUSINESS STUDIES"],
    "Agriculture": ["Agriculture", "agriculture "],
    "Technology": ["Technology", "TECHNOLOGY"],
    "Medicine": ["Medicine", "medicine"],
}
EXAM_FEE = {"Applied Sciences": 2000, "Technology": 2000, "Medicine": 2000,
            "Business Studies": 1500, "Agriculture": 1500}
REG_FEE, LIBRARY_FEE, LAB_FEE, HOSTEL_FEE = 2500, 1000, 3000, 18000

SEMESTERS = [("2025 S1", date(2025, 1, 13)), ("2025 S2", date(2025, 7, 14)),
             ("2026 S1", date(2026, 1, 12)), ("2026 S2", date(2026, 7, 13))]
BATCH_SEMESTERS = {2022: [0, 1], 2023: [0, 1, 2, 3], 2024: [0, 1, 2, 3], 2025: [1, 2, 3]}
DUE_OFFSET = {"Registration": 0, "Hostel": 14, "Lab": 60, "Library": 90, "Exam": 126}

METHOD_VARIANTS = {
    "Online": ["Online", "online", "ONLINE "],
    "Bank Deposit": ["Bank Deposit", "bank deposit", "Bank deposit "],
    "Cash": ["Cash", "cash ", " CASH"],
}
CLERKS = ["K. Perera", "N. Silva", "R. Fernando", "S. Jayasinghe"]
ONLINE_RECORDER = "Online Portal"

FIRST = ["Nimal", "Kasun", "Tharindu", "Dilshan", "Chamara", "Isuru", "Sahan", "Ruwan", "Pasindu",
         "Hasitha", "Lahiru", "Supun", "Janith", "Malith", "Ravindu", "Sachin", "Dineth", "Kavindu",
         "Sanduni", "Ishara", "Nadeesha", "Dilini", "Hiruni", "Kavindya", "Sachini", "Tharushi",
         "Nethmi", "Piumi", "Chathurika", "Madushani", "Anjali", "Yashodha", "Imesha", "Sewwandi",
         "Hansika", "Gayani", "Shehani", "Rashmi", "Thilini", "Oshadi"]
LAST = ["Perera", "Fernando", "Silva", "Jayasinghe", "Wijesinghe", "Bandara", "Rathnayake",
        "Gunawardena", "Dissanayake", "Herath", "Karunaratne", "Wickramasinghe", "Senanayake",
        "Abeysekara", "Weerasinghe", "Ekanayake", "Kumara", "Samarasinghe", "Liyanage", "Rajapaksha"]

rng = random.Random(SEED)


# ---- 1. Students: search attributes until fee count and billed hit target ---
def fee_types_for(fac: str, hostel: bool) -> list[tuple[str, int]]:
    out = [("Registration", REG_FEE), ("Exam", EXAM_FEE[fac]), ("Library", LIBRARY_FEE)]
    if fac != "Business Studies":
        out.append(("Lab", LAB_FEE))
    if hostel:
        out.append(("Hostel", HOSTEL_FEE))
    return out


def student_totals(s) -> tuple[int, int]:
    n = len(BATCH_SEMESTERS[s["Batch"]])
    ft = fee_types_for(s["Faculty"], s["Hostel"])
    return n * len(ft), n * sum(a for _, a in ft)


def make_students():
    students = []
    for i in range(N_STUDENTS):
        students.append({
            "Faculty": rng.choices(FACULTIES, FACULTY_WEIGHTS)[0],
            "Batch": rng.choice(list(BATCH_SEMESTERS)),
            "Hostel": rng.random() < 0.36,
        })

    import math
    C = sum(student_totals(s)[0] for s in students)
    B = sum(student_totals(s)[1] for s in students)

    def loss(c, b):
        return abs(c - N_FEES) * 4000 + abs(b - TOTAL_BILLED)

    cur, T = loss(C, B), 20000.0
    for _ in range(400000):
        if cur == 0:
            break
        s = rng.choice(students)
        old = dict(s)
        c0, b0 = student_totals(s)
        k = rng.random()
        if k < 0.4:
            s["Hostel"] = not s["Hostel"]
        elif k < 0.6:
            s["Batch"] = rng.choice(list(BATCH_SEMESTERS))
        else:
            s["Faculty"] = rng.choices(FACULTIES, FACULTY_WEIGHTS)[0]
        c1, b1 = student_totals(s)
        nc, nb = C - c0 + c1, B - b0 + b1
        new = loss(nc, nb)
        if new <= cur or rng.random() < math.exp((cur - new) / T):
            C, B, cur = nc, nb, new
        else:
            s.update(old)
        T = max(50.0, T * 0.99997)
    assert cur == 0, f"could not hit fee targets, loss={cur}"

    # Sort for tidy IDs: by batch then faculty
    students.sort(key=lambda s: (s["Batch"], FACULTIES.index(s["Faculty"])))
    used = set()
    for i, s in enumerate(students, 1):
        s["StudentID"] = f"STU{i:04d}"
        while True:
            name = f"{rng.choice(FIRST)} {rng.choice(LAST)}"
            if name not in used:
                used.add(name)
                break
        s["StudentName"] = name
    return students


students = make_students()
stu_by_id = {s["StudentID"]: s for s in students}

# ---- 2. Fees ---------------------------------------------------------------
fees = []
for s in students:
    for si in BATCH_SEMESTERS[s["Batch"]]:
        sem, start = SEMESTERS[si]
        for ftype, amt in fee_types_for(s["Faculty"], s["Hostel"]):
            fees.append({"StudentID": s["StudentID"], "Semester": sem, "SemIdx": si,
                         "FeeType": ftype, "Amount": amt,
                         "DueDate": start + timedelta(days=DUE_OFFSET[ftype])})
fees.sort(key=lambda f: (f["SemIdx"], f["StudentID"], list(DUE_OFFSET).index(f["FeeType"])))
for i, f in enumerate(fees, 1):
    f["FeeID"] = f"F{i:05d}"
    f["role"] = "paid"           # default: paid in full by one normal payment
assert len(fees) == N_FEES and sum(f["Amount"] for f in fees) == TOTAL_BILLED
assert all(f["DueDate"] != REPORT_DATE for f in fees)

# ---- 3. Not-yet-due unpaid exam fees (2026 S2, due 16 Nov 2026) -------------
for amt, n in NOT_YET_DUE_UNPAID.items():
    pool = [f for f in fees if f["FeeType"] == "Exam" and f["SemIdx"] == 3 and f["Amount"] == amt]
    for f in rng.sample(pool, n):
        f["role"] = "unpaid_notdue"

# ---- 4. Exception fees -----------------------------------------------------
def take(amount, n, ftype=None, before=REPORT_DATE):
    pool = [f for f in fees if f["role"] == "paid" and f["Amount"] == amount
            and f["DueDate"] < before and (ftype is None or f["FeeType"] == ftype)]
    return rng.sample(pool, n)

exceptions = {
    "duplicate": [f for a, n in [(2500, 5), (3000, 3), (2000, 1)] for f in take(a, n)],
    "overpay": [f for a, n in [(18000, 2), (3000, 1), (2500, 1), (1000, 1)] for f in take(a, n)],
    "noreceipt": [f for a, n in [(2500, 2), (1500, 4), (1000, 3), (3000, 1)] for f in take(a, n)],
    "bigcash": take(18000, 7, "Hostel"),
    "future": [f for a, n in [(2500, 2), (1500, 2)] for f in take(a, n, before=date(2026, 6, 1))],
}
for role, fl in exceptions.items():
    for f in fl:
        assert f["role"] == "paid"
        f["role"] = role

# ---- 5. Overdue fees: 364 fees, 189 students, LKR 1,101,600 ----------------
SEM_WEIGHT = {0: 1.0, 1: 1.6, 2: 3.0, 3: 6.0}
TYPE_WEIGHT = {"Registration": 1.0, "Exam": 1.0, "Library": 1.3, "Lab": 1.0, "Hostel": 0.35}


def pick_overdue():
    cands = {}
    for f in fees:
        if f["role"] == "paid" and f["DueDate"] < REPORT_DATE:
            cands.setdefault(f["StudentID"], []).append(f)
    sids = rng.sample(sorted(cands), STUDENTS_OVERDUE)
    chosen = []
    for sid in sids:
        fl = cands[sid]
        w = [SEM_WEIGHT[f["SemIdx"]] * TYPE_WEIGHT[f["FeeType"]] for f in fl]
        chosen.append(rng.choices(fl, w)[0])
    remaining = [f for sid in sids for f in cands[sid] if f not in chosen]
    w = [SEM_WEIGHT[f["SemIdx"]] * TYPE_WEIGHT[f["FeeType"]] for f in remaining]
    extra = set()
    while len(extra) < OVERDUE_FEES - STUDENTS_OVERDUE:
        extra.add(id(rng.choices(remaining, w)[0]))
    chosen += [f for f in remaining if id(f) in extra]
    return chosen


for attempt in range(5000):
    od = pick_overdue()
    # larger fees make the best partial payments (more room to tune)
    big = [f for f in od if f["Amount"] >= 2500]
    if len(big) < OVERDUE_PARTIAL:
        continue
    partial = rng.sample(big, OVERDUE_PARTIAL)
    pids = {id(f) for f in partial}
    unpaid = [f for f in od if id(f) not in pids]
    U = sum(f["Amount"] for f in unpaid)
    R = OVERDUE_AMOUNT - U                      # balance that partial fees must carry
    lo, hi = 100 * OVERDUE_PARTIAL, sum(f["Amount"] - 100 for f in partial)
    if lo <= R <= hi:
        break
else:
    raise RuntimeError("no feasible overdue selection")

for f in unpaid:
    f["role"] = "unpaid_overdue"
# spread R over the partial fees in multiples of 100
bal = [100] * len(partial)
left = R - sum(bal)
order = list(range(len(partial)))
while left > 0:
    rng.shuffle(order)
    for i in order:
        room = partial[i]["Amount"] - 100 - bal[i]
        if room <= 0 or left <= 0:
            continue
        step = min(room, left, 100 * rng.randint(1, max(1, room // 300)))
        bal[i] += step
        left -= step
for f, b in zip(partial, bal):
    f["role"] = "partial"
    f["Balance"] = b

# ---- 6. Split payments so the payment count lands on 3,137 -----------------
z = sum(1 for f in fees if f["role"] in ("unpaid_notdue", "unpaid_overdue"))
n_split = N_PAYMENTS - 9 - (N_FEES - z)    # 9 = extra rows from duplicate pairs
assert n_split >= 0, n_split
SPLITS = {18000: [(10000, 8000), (12000, 6000), (9500, 8500)], 3000: [(2000, 1000), (1800, 1200)],
          2500: [(1500, 1000), (2000, 500)]}
split_pool = [f for f in fees if f["role"] == "paid" and f["Amount"] in SPLITS]
for f in rng.sample(split_pool, n_split):
    f["role"] = "split"

# ---- 7. Payments -----------------------------------------------------------
payments = []


def pay_date(f, early=-15, late=45):
    start = SEMESTERS[f["SemIdx"]][1]
    d = f["DueDate"] + timedelta(days=rng.randint(early, late))
    d = max(d, start - timedelta(days=7))
    return min(d, REPORT_DATE - timedelta(days=rng.randint(1, 20)))


def method_for(amount, allow_cash=True):
    m = rng.choices(["Online", "Bank Deposit", "Cash"], [50, 35, 15])[0]
    if m == "Cash" and (amount > 15000 or not allow_cash):
        m = rng.choice(["Online", "Bank Deposit"])
    return m


def add(f, amount, d, method, receipt=True, by=None):
    if by is None:
        by = ONLINE_RECORDER if method == "Online" else rng.choice(CLERKS)
    payments.append({"FeeID": f["FeeID"], "PaymentDate": d, "AmountPaid": amount,
                     "Method": method, "has_receipt": receipt, "RecordedBy": by})


for f in fees:
    r, a = f["role"], f["Amount"]
    if r in ("unpaid_notdue", "unpaid_overdue"):
        continue
    if r == "paid":
        if f["FeeType"] == "Exam" and f["SemIdx"] == 3:     # paid early, not yet due
            d = date(2026, 8, 1) + timedelta(days=rng.randint(0, 85))
        else:
            d = pay_date(f)
        add(f, a, d, method_for(a))
    elif r == "split":
        a1, a2 = rng.choice(SPLITS[a])
        d1 = pay_date(f, -10, 10)
        d2 = min(d1 + timedelta(days=rng.randint(20, 60)), REPORT_DATE - timedelta(days=1))
        add(f, a1, d1, method_for(a1))
        add(f, a2, d2, method_for(a2))
    elif r == "partial":
        amt = a - f["Balance"]
        add(f, amt, pay_date(f, -5, 20), method_for(amt))
    elif r == "duplicate":
        d = pay_date(f)
        m = rng.choice(["Bank Deposit", "Cash", "Online"])
        by = "K. Perera" if rng.random() < 0.7 and m != "Online" else None
        add(f, a, d, m, by=by)
        add(f, a, d + timedelta(days=rng.choice([0, 0, 1])), m, by=by)
    elif r == "overpay":
        # extra-zero typing errors; exactly one is also a large cash payment
        m = "Cash" if f is exceptions["overpay"][0] else "Bank Deposit"
        add(f, a * 10, pay_date(f), m, by=rng.choice(["K. Perera", "K. Perera", "S. Jayasinghe"]))
    elif r == "noreceipt":
        m = rng.choice(["Cash", "Bank Deposit"])
        add(f, a, pay_date(f), m, receipt=False,
            by=rng.choice(["N. Silva", "N. Silva", "K. Perera", "R. Fernando"]))
    elif r == "bigcash":
        add(f, a, pay_date(f), "Cash", by=rng.choice(CLERKS))
    elif r == "future":
        real = pay_date(f)
        typo = real.replace(year=real.year + 1)            # year typed wrongly
        if typo <= REPORT_DATE:
            typo = REPORT_DATE + timedelta(days=rng.randint(3, 60))
        add(f, a, typo, rng.choice(["Cash", "Bank Deposit"]), by="R. Fernando")

assert len(payments) == N_PAYMENTS, len(payments)
payments.sort(key=lambda p: (p["PaymentDate"], p["FeeID"]))
rc = 1000
for i, p in enumerate(payments, 1):
    p["PaymentID"] = f"P{i:05d}"
    if p["has_receipt"]:
        rc += 1
        p["ReceiptNo"] = f"RC{p['PaymentDate'].year}-{rc:05d}"
    else:
        p["ReceiptNo"] = None

# ---- 8. Messy text ---------------------------------------------------------
for s in students:
    s["FacultyRaw"] = rng.choice(FACULTY_VARIANTS[s["Faculty"]])
# make sure every one of the 12 variants appears at least once
for fac, variants in FACULTY_VARIANTS.items():
    group = [s for s in students if s["Faculty"] == fac]
    for v, s in zip(variants, rng.sample(group, len(variants))):
        s["FacultyRaw"] = v
for p in payments:
    clean = p["Method"]
    p["MethodRaw"] = clean if rng.random() < 0.8 else rng.choice(METHOD_VARIANTS[clean])

# ---- 9. Write workbook -----------------------------------------------------
students_df = pd.DataFrame([{
    "StudentID": s["StudentID"], "StudentName": s["StudentName"], "Faculty": s["FacultyRaw"],
    "Batch": str(s["Batch"]), "Hostel": "Yes" if s["Hostel"] else "No"} for s in students])
fees_df = pd.DataFrame([{
    "FeeID": f["FeeID"], "StudentID": f["StudentID"], "Semester": f["Semester"],
    "FeeType": f["FeeType"], "Amount": f["Amount"], "DueDate": f["DueDate"]} for f in fees])
pay_df = pd.DataFrame([{
    "PaymentID": p["PaymentID"], "FeeID": p["FeeID"], "PaymentDate": p["PaymentDate"],
    "AmountPaid": p["AmountPaid"], "Method": p["MethodRaw"], "ReceiptNo": p["ReceiptNo"],
    "RecordedBy": p["RecordedBy"]} for p in payments])

about = [
    ["Student Dues Monitoring – practice dataset"],
    [""],
    ["All data is fictional. Snapshot as at 31 Oct 2026. Amounts in LKR."],
    ["Sheets: Students (one row per student), Fees (one row per fee charged per semester), "
     "Payments (one row per payment received)."],
    [""],
    ["Fee schedule (per semester)"],
    ["Registration – LKR 2,500, due on the semester start date"],
    ["Hostel – LKR 18,000 (hostel residents), due 14 days after semester start"],
    ["Lab – LKR 3,000 (all faculties except Business Studies), due 60 days after semester start"],
    ["Library – LKR 1,000, due 90 days after semester start"],
    ["Exam – LKR 2,000 (Applied Sciences, Technology, Medicine) or LKR 1,500 (Business Studies, "
     "Agriculture), due 126 days after semester start"],
    [""],
    ["Payment policy rules to test (see Fee_Payment_Policy.pdf)"],
    ["Rule 1 – Duplicates: the same fee should not be paid twice for the same amount."],
    ["Rule 2 – Overpayments: a single payment must not exceed the fee it pays."],
    ["Rule 3 – Future-dated entries: payments must not be dated after the report date (31 Oct 2026)."],
    ["Rule 4 – Receipts: every payment must have a receipt number."],
    ["Rule 5 – Large cash: cash payments over LKR 15,000 need prior Bursar approval and are flagged."],
    [""],
    ["Known data-quality issues: Faculty and Method contain inconsistent spacing and capitalisation; "
     "clean them in Power Query (Trim, then Capitalize Each Word). Do not delete blank ReceiptNo values."],
]

OUT.parent.mkdir(parents=True, exist_ok=True)
with pd.ExcelWriter(OUT, engine="openpyxl") as xw:
    pd.DataFrame(about).to_excel(xw, sheet_name="About", index=False, header=False)
    students_df.to_excel(xw, sheet_name="Students", index=False)
    fees_df.to_excel(xw, sheet_name="Fees", index=False)
    pay_df.to_excel(xw, sheet_name="Payments", index=False)

wb = load_workbook(OUT)
head_fill = PatternFill("solid", fgColor="1F3A5F")
for ws in wb.worksheets:
    if ws.title == "About":
        ws.column_dimensions["A"].width = 120
        ws["A1"].font = Font(bold=True, size=14)
        for r in (6, 13):
            ws.cell(r, 1).font = Font(bold=True)
        for row in ws.iter_rows():
            for c in row:
                c.alignment = Alignment(wrap_text=True, vertical="top")
        continue
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = head_fill
    ws.freeze_panes = "A2"
    for i, col in enumerate(ws.iter_cols(min_row=1, max_row=1), 1):
        ws.column_dimensions[get_column_letter(i)].width = max(12, len(str(col[0].value)) + 4)
    for col_name in ("DueDate", "PaymentDate"):
        for c in ws[1]:
            if c.value == col_name:
                for cell in ws[get_column_letter(c.column)][1:]:
                    cell.number_format = "yyyy-mm-dd"
wb.save(OUT)
print(f"Wrote {OUT}  students={len(students_df)} fees={len(fees_df)} payments={len(pay_df)}")
