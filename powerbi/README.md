# Power BI build

Power BI Desktop is free and Windows-only. Build the report from the files in this folder, then save it here as **`Student_Dues_Monitoring.pbix`**.

| File | Use it for |
|---|---|
| `power-query/DataFile.pq` | Parameter holding the path to `data/Student_Dues_Data.xlsx` |
| `power-query/Students.pq`, `Fees.pq`, `Payments.pq` | Cleaning steps (Advanced Editor). The same steps can be done by clicking: Trim → Capitalize Each Word → set types |
| `dax/calculated-columns.dax` | 5 columns on Fees, 1 on Payments |
| `dax/measures.dax` | 9 measures in a `_Measures` table, plus an optional dynamic title |
| `theme/student-dues-theme.json` | View → Themes → Browse for themes. Blue = collected, orange = overdue/exception |

## 1. Load and clean (45 min)

1. **Home → Get data → Excel workbook** → pick `data/Student_Dues_Data.xlsx`. Tick Students, Fees and Payments (not About) → **Transform Data**.
2. **View** → tick **Column quality** and **Column distribution**.
3. **Students**: Faculty → Trim → Capitalize Each Word. There should be 5 faculties (from 12 spellings). Batch → Whole Number.
4. **Fees**: Amount → Whole Number, DueDate → **Date** (not Date/Time).
5. **Payments**: Method → Trim → Capitalize Each Word. There should be 3 values (Bank Deposit, Cash, Online). PaymentDate → Date, AmountPaid → Whole Number. Leave blank ReceiptNo values alone.
6. **Close & Apply** and save as `Student_Dues_Monitoring.pbix`.

## 2. Model (15 min)

| From (one) | To (many) | Cardinality | Cross-filter |
|---|---|---|---|
| Students[StudentID] | Fees[StudentID] | 1:* | Single |
| Fees[FeeID] | Payments[FeeID] | 1:* | Single |

## 3. DAX (1–1.5 h)

Type the formulas from `dax/` yourself. Then put every measure on a Card with no filters and compare against the check table in the main README. If something is off, check the Date types (step 1) and the relationships (step 2).

## 4. Report pages

**Page 1 – Collections overview**

| Visual | Fields |
|---|---|
| 4 Cards | Total Billed, Total Collected, Collection Rate, Outstanding |
| Clustered column | X: Fees[Semester] · Y: Total Billed, Total Collected |
| Clustered bar | Y: Students[Faculty] · X: Collection Rate (sort ascending) |
| Slicers | Students[Faculty], Students[Batch], Fees[FeeType] |

**Page 2 – Overdue & ageing**

| Visual | Fields |
|---|---|
| 3 Cards | Overdue Amount, Overdue Fees, Students Overdue |
| Column | X: Fees[Ageing] · Y: Overdue Amount |
| Matrix | Rows: Students[Faculty] · Columns: Fees[FeeType] · Values: Overdue Amount (conditional formatting → background colour) |
| Table | StudentID, StudentName, Faculty, Overdue Amount · filter StudentID Top N = 10 by Overdue Amount |

Copy the slicers from Page 1 and choose **Sync**.

**Page 3 – Exceptions**

| Visual | Fields |
|---|---|
| 2 Cards | Exceptions, Exception Value |
| Bar | Y: Payments[Exception] · X: Exceptions · visual filter Exception is not "None" |
| Bar | Y: Payments[RecordedBy] · X: Exceptions |
| Table (exception register) | PaymentID, FeeID, Students[StudentName], PaymentDate, AmountPaid, Method, ReceiptNo, RecordedBy, Exception · same filter |

Design rules: same two colours on every page, a takeaway title on each page ("LKR 1.1M overdue across 189 students"), **Format → Align** everything, plain-English visual titles.

## 5. Export

- **File → Export → Export to PDF** → save as `docs/Student_Dues_Monitoring.pdf`.
- Screenshot each page into `docs/screenshots/` as `page1-overview.png`, `page2-overdue.png`, `page3-exceptions.png` (the main README already links these names).
