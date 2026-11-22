"""
Re-implement the Power Query cleaning and the DAX logic in pandas and check
the results against the expected figures in the README.

Use it to confirm the dataset is intact, and as an independent cross-check
of the numbers the Power BI report shows.

    python scripts/validate_numbers.py          # prints the check table, exits 1 on mismatch
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

DATA = Path(__file__).resolve().parents[1] / "data" / "Student_Dues_Data.xlsx"
REPORT_DATE = pd.Timestamp(2026, 10, 31)

EXPECTED = {
    "Total Billed": 12_359_000,
    "Total Collected": 11_459_400,
    "Collection Rate": "92.7%",
    "Outstanding": 1_305_600,
    "Overdue Amount": 1_101_600,
    "Overdue Fees": 364,
    "Students Overdue": 189,
    "Exceptions": 44,
    "Exception Value": 623_000,
}
EXPECTED_BY_TYPE = {
    "Possible duplicate": (18, 47_000),
    "Missing receipt": (10, 17_000),
    "Large cash payment": (7, 126_000),
    "Overpayment": (5, 425_000),
    "Future-dated": (4, 8_000),
}


def load():
    xl = pd.read_excel(DATA, sheet_name=None)
    st, fe, pa = xl["Students"], xl["Fees"], xl["Payments"]
    # --- Power Query step equivalents ---
    st["Faculty"] = st["Faculty"].str.strip().str.title()      # Trim + Capitalize Each Word
    st["Batch"] = st["Batch"].astype(int)
    fe["Amount"] = fe["Amount"].astype(int)
    fe["DueDate"] = pd.to_datetime(fe["DueDate"]).dt.normalize()
    pa["Method"] = pa["Method"].str.strip().str.title()
    pa["PaymentDate"] = pd.to_datetime(pa["PaymentDate"]).dt.normalize()
    pa["AmountPaid"] = pa["AmountPaid"].astype(int)
    return st, fe, pa


def compute():
    """Return cleaned Students, Fees (with calculated columns) and Payments (with Exception)."""
    st, fe, pa = load()

    # --- Fees calculated columns ---
    paid = pa.groupby("FeeID")["AmountPaid"].sum()
    fe["Paid"] = fe["FeeID"].map(paid).fillna(0).astype(int)
    fe["Balance"] = fe["Amount"] - fe["Paid"]

    def status(r):
        if r.Balance < 0:
            return "Overpaid"
        if r.Balance == 0:
            return "Paid"
        if r.DueDate < REPORT_DATE:
            return "Overdue"
        return "Not yet due"

    fe["Status"] = fe.apply(status, axis=1)
    fe["DaysOverdue"] = (REPORT_DATE - fe["DueDate"]).dt.days.where(fe["Status"] == "Overdue")

    # --- Payments exception column (first rule broken wins) ---
    pa = pa.merge(fe[["FeeID", "Amount"]], on="FeeID", how="left")
    dup = pa.groupby(["FeeID", "AmountPaid"])["PaymentID"].transform("count") > 1

    def exception(i, r):
        if dup.iloc[i]:
            return "Possible duplicate"
        if r.AmountPaid > r.Amount:
            return "Overpayment"
        if r.PaymentDate > REPORT_DATE:
            return "Future-dated"
        if pd.isna(r.ReceiptNo) or str(r.ReceiptNo).strip() == "":
            return "Missing receipt"
        if r.Method == "Cash" and r.AmountPaid > 15000:
            return "Large cash payment"
        return "None"

    pa["Exception"] = [exception(i, r) for i, r in enumerate(pa.itertuples())]
    return st, fe, pa


def main() -> int:
    st, fe, pa = compute()
    problems = []
    if st["Faculty"].nunique() != 5:
        problems.append(f"Faculty values after cleaning: {sorted(st['Faculty'].unique())}")
    if sorted(pa["Method"].unique()) != ["Bank Deposit", "Cash", "Online"]:
        problems.append(f"Method values after cleaning: {sorted(pa['Method'].unique())}")
    flagged = pa[pa["Exception"] != "None"]

    # --- Measures ---
    billed, collected = fe["Amount"].sum(), pa["AmountPaid"].sum()
    od = fe[fe["Status"] == "Overdue"]
    actual = {
        "Total Billed": int(billed),
        "Total Collected": int(collected),
        "Collection Rate": f"{collected / billed:.1%}",
        "Outstanding": int(fe.loc[fe["Balance"] > 0, "Balance"].sum()),
        "Overdue Amount": int(od["Balance"].sum()),
        "Overdue Fees": len(od),
        "Students Overdue": od["StudentID"].nunique(),
        "Exceptions": len(flagged),
        "Exception Value": int(flagged["AmountPaid"].sum()),
    }

    print(f"Rows: Students={len(st)}  Fees={len(fe)}  Payments={len(pa)}\n")
    print(f"{'Measure':<20}{'Expected':>14}{'Actual':>14}")
    for k, exp in EXPECTED.items():
        got = actual[k]
        ok = "OK" if got == exp else "MISMATCH"
        print(f"{k:<20}{exp:>14,}{got:>14,}  {ok}" if isinstance(exp, int)
              else f"{k:<20}{exp:>14}{got:>14}  {ok}")
        if got != exp:
            problems.append(f"{k}: expected {exp}, got {got}")

    print(f"\n{'Exception':<22}{'Payments':>10}{'Value (LKR)':>14}")
    by_type = flagged.groupby("Exception")["AmountPaid"].agg(["count", "sum"])
    for k, (n, v) in EXPECTED_BY_TYPE.items():
        gn, gv = (int(by_type.loc[k, "count"]), int(by_type.loc[k, "sum"])) if k in by_type.index else (0, 0)
        flag = "" if (gn, gv) == (n, v) else "  MISMATCH"
        print(f"{k:<22}{gn:>10}{gv:>14,}{flag}")
        if (gn, gv) != (n, v):
            problems.append(f"{k}: expected {(n, v)}, got {(gn, gv)}")

    big_cash_all = ((pa["Method"] == "Cash") & (pa["AmountPaid"] > 15000)).sum()
    overpaid = (fe["Status"] == "Overpaid").sum()
    print(f"\nPayments breaking the large-cash rule (any label): {big_cash_all}  (expected 8)")
    print(f"Fees with status Overpaid: {overpaid}  (expected 14)")
    if big_cash_all != 8:
        problems.append("large-cash rule count")
    if overpaid != 14:
        problems.append("overpaid fee count")

    ageing = pd.cut(od["DaysOverdue"], [0, 30, 90, 180, 10_000],
                    labels=["0–30 days", "31–90 days", "91–180 days", "Over 180 days"])
    print("\nOverdue by ageing bucket (LKR):")
    print(od.groupby(ageing, observed=False)["Balance"].sum().to_string())

    if problems:
        print("\nFAILED:\n  " + "\n  ".join(problems))
        return 1
    print("\nAll checks passed.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
