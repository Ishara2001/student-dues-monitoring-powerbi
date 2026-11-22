"""
Render static previews of the three dashboard pages from the same logic as the DAX
(scripts/validate_numbers.py). These are a reference for layout and expected values;
replace them with real Power BI screenshots in docs/screenshots/.
"""
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

from validate_numbers import compute

ROOT = Path(__file__).resolve().parents[1]
OUT = ROOT / "previews"
BLUE, ORANGE, GREY, INK, MUTED = "#1F6FB2", "#E4572E", "#C8D1DA", "#1F2933", "#52606D"
plt.rcParams.update({"font.family": "DejaVu Sans", "axes.spines.top": False,
                     "axes.spines.right": False, "axes.edgecolor": GREY,
                     "axes.labelcolor": MUTED, "xtick.color": MUTED, "ytick.color": MUTED})

st, fe, pa = compute()
fe = fe.merge(st[["StudentID", "StudentName", "Faculty"]], on="StudentID")
pa = pa.merge(fe[["FeeID", "StudentID", "StudentName", "Faculty"]], on="FeeID")


def lkr(v):
    return f"{v:,.0f}"


def page(title, subtitle):
    fig = plt.figure(figsize=(16, 9), dpi=110)
    fig.patch.set_facecolor("white")
    fig.text(0.03, 0.95, title, fontsize=22, weight="bold", color=INK)
    fig.text(0.03, 0.915, subtitle, fontsize=11, color=MUTED)
    fig.text(0.97, 0.015, "Python preview of the Power BI layout · data as at 31 Oct 2026 · LKR",
             fontsize=8, color=MUTED, ha="right")
    return fig


def cards(fig, items, y=0.75, h=0.12):
    n = len(items)
    w = 0.94 / n
    for i, (label, value, color) in enumerate(items):
        ax = fig.add_axes([0.03 + i * w, y, w - 0.015, h])
        ax.set_xticks([]), ax.set_yticks([])
        for s in ax.spines.values():
            s.set_visible(True)
            s.set_color(GREY)
        ax.axvline(0, color=color, lw=6)
        ax.text(0.06, 0.62, value, fontsize=24, weight="bold", color=INK, transform=ax.transAxes)
        ax.text(0.06, 0.22, label, fontsize=11, color=MUTED, transform=ax.transAxes)


# ---- Page 1 ----------------------------------------------------------------
billed, collected = fe["Amount"].sum(), pa["AmountPaid"].sum()
outstanding = fe.loc[fe["Balance"] > 0, "Balance"].sum()
fig = page(f"{collected / billed:.1%} of fees collected",
           "Collections overview – all faculties, batches and fee types")
cards(fig, [("Total Billed", lkr(billed), BLUE), ("Total Collected", lkr(collected), BLUE),
            ("Collection Rate", f"{collected / billed:.1%}", BLUE),
            ("Outstanding", lkr(outstanding), ORANGE)])
ax = fig.add_axes([0.05, 0.08, 0.45, 0.58])
sem = fe.groupby("Semester")["Amount"].sum()
col = pa.merge(fe[["FeeID", "Semester"]], on="FeeID").groupby("Semester")["AmountPaid"].sum()
x = range(len(sem))
ax.bar([i - 0.2 for i in x], sem.values / 1e6, 0.4, color=GREY, label="Billed")
ax.bar([i + 0.2 for i in x], col.reindex(sem.index).values / 1e6, 0.4, color=BLUE, label="Collected")
ax.set_xticks(list(x), sem.index)
ax.set_ylabel("LKR million")
ax.set_title("Billed vs collected by semester", loc="left", fontsize=13, color=INK)
ax.legend(frameon=False)
ax = fig.add_axes([0.62, 0.08, 0.34, 0.58])
fb = fe.groupby("Faculty")["Amount"].sum()
fc = pa.groupby("Faculty")["AmountPaid"].sum()
rate = (fc / fb).sort_values(ascending=False)
ax.barh(rate.index, rate.values * 100, color=BLUE)
for i, v in enumerate(rate.values):
    ax.text(v * 100 + 0.3, i, f"{v:.1%}", va="center", fontsize=10, color=INK)
ax.set_xlim(80, 100)
ax.set_title("Collection rate by faculty (weakest at the bottom)", loc="left", fontsize=13, color=INK)
fig.savefig(OUT / "page1-overview.png", bbox_inches="tight")
plt.close(fig)

# ---- Page 2 ----------------------------------------------------------------
od = fe[fe["Status"] == "Overdue"].copy()
od["Ageing"] = pd.cut(od["DaysOverdue"], [0, 30, 90, 180, 10_000],
                      labels=["0–30 days", "31–90 days", "91–180 days", "Over 180 days"])
fig = page(f"LKR {od['Balance'].sum() / 1e6:.1f}M overdue across {od['StudentID'].nunique()} students",
           "Overdue & ageing")
cards(fig, [("Overdue Amount", lkr(od["Balance"].sum()), ORANGE),
            ("Overdue Fees", lkr(len(od)), ORANGE),
            ("Students Overdue", lkr(od["StudentID"].nunique()), ORANGE)])
ax = fig.add_axes([0.05, 0.40, 0.40, 0.28])
ag = od.groupby("Ageing", observed=False)["Balance"].sum()
ax.bar(ag.index.astype(str), ag.values / 1000, color=ORANGE)
ax.set_ylabel("LKR thousand")
ax.set_title("Overdue amount by days overdue", loc="left", fontsize=13, color=INK)
ax = fig.add_axes([0.05, 0.06, 0.40, 0.26])
mx = od.pivot_table(index="Faculty", columns="FeeType", values="Balance", aggfunc="sum", fill_value=0)
ax.imshow(mx.values, cmap="Oranges", aspect="auto")
ax.set_xticks(range(len(mx.columns)), mx.columns)
ax.set_yticks(range(len(mx.index)), mx.index)
for i in range(mx.shape[0]):
    for j in range(mx.shape[1]):
        v = mx.values[i, j]
        ax.text(j, i, lkr(v), ha="center", va="center", fontsize=9,
                color="white" if v > 0.55 * mx.values.max() else INK)
ax.set_title("Overdue by faculty and fee type", loc="left", fontsize=13, color=INK)
ax.spines[:].set_visible(False)
ax = fig.add_axes([0.52, 0.06, 0.45, 0.62])
ax.axis("off")
top = (od.groupby(["StudentID", "StudentName", "Faculty"])["Balance"].sum()
       .sort_values(ascending=False).head(10).reset_index())
tbl = ax.table(cellText=[[r.StudentID, r.StudentName, r.Faculty, lkr(r.Balance)] for r in top.itertuples()],
               colLabels=["StudentID", "Student", "Faculty", "Overdue (LKR)"], loc="upper center",
               cellLoc="left", colWidths=[0.17, 0.33, 0.3, 0.2])
tbl.auto_set_font_size(False)
tbl.set_fontsize(10)
tbl.scale(1, 1.7)
for (r, c), cell in tbl.get_celld().items():
    cell.set_edgecolor(GREY)
    if r == 0:
        cell.set_facecolor("#1F3A5F")
        cell.get_text().set_color("white")
ax.set_title("Top 10 overdue students", loc="left", fontsize=13, color=INK)
fig.savefig(OUT / "page2-overdue.png", bbox_inches="tight")
plt.close(fig)

# ---- Page 3 ----------------------------------------------------------------
ex = pa[pa["Exception"] != "None"]
fig = page(f"{len(ex)} of {len(pa):,} payments broke a policy rule (LKR {ex['AmountPaid'].sum():,.0f})",
           "Exceptions – five rules tested on every payment")
cards(fig, [("Exceptions", lkr(len(ex)), ORANGE), ("Exception Value", lkr(ex["AmountPaid"].sum()), ORANGE)],
      h=0.12)
ax = fig.add_axes([0.13, 0.40, 0.33, 0.28])
bt = ex["Exception"].value_counts().sort_values()
ax.barh(bt.index, bt.values, color=ORANGE)
for i, v in enumerate(bt.values):
    ax.text(v + 0.2, i, str(v), va="center", fontsize=10)
ax.set_title("Exceptions by rule", loc="left", fontsize=13, color=INK)
ax = fig.add_axes([0.62, 0.40, 0.33, 0.28])
br = ex["RecordedBy"].value_counts().sort_values()
ax.barh(br.index, br.values, color=ORANGE)
for i, v in enumerate(br.values):
    ax.text(v + 0.2, i, str(v), va="center", fontsize=10)
ax.set_title("Exceptions by who recorded the payment", loc="left", fontsize=13, color=INK)
ax = fig.add_axes([0.03, 0.02, 0.94, 0.32])
ax.axis("off")
reg = ex.sort_values("AmountPaid", ascending=False).head(8)
tbl = ax.table(cellText=[[r.PaymentID, r.FeeID, r.StudentName, r.PaymentDate.strftime("%Y-%m-%d"),
                          lkr(r.AmountPaid), r.Method, "" if pd.isna(r.ReceiptNo) else r.ReceiptNo,
                          r.RecordedBy, r.Exception] for r in reg.itertuples()],
               colLabels=["PaymentID", "FeeID", "Student", "Date", "Amount", "Method", "ReceiptNo",
                          "RecordedBy", "Exception"], loc="upper center", cellLoc="left")
tbl.auto_set_font_size(False)
tbl.set_fontsize(9)
tbl.scale(1, 1.45)
for (r, c), cell in tbl.get_celld().items():
    cell.set_edgecolor(GREY)
    if r == 0:
        cell.set_facecolor("#1F3A5F")
        cell.get_text().set_color("white")
ax.set_title("Exception register (largest 8 of 44)", loc="left", fontsize=13, color=INK)
fig.savefig(OUT / "page3-exceptions.png", bbox_inches="tight")
plt.close(fig)
print("Previews written to", OUT)
