# Weekly payment exception alert (Power Automate)

A scheduled cloud flow that emails the exception register to a reviewer every Monday at 9:00, and only when there is something to report.

**Account:** sign in at [make.powerautomate.com](https://make.powerautomate.com) with a work or school (university) account. Both connectors used below — Excel Online (Business) and Office 365 Outlook — are standard connectors. If the account has no access, Power Automate for desktop (free on Windows) can do the same job from a local Excel file.

## Flow design

```
Recurrence (every 1 week, Monday 09:00, Asia/Colombo)
  └─ Excel Online (Business): List rows present in a table
       Location: OneDrive for Business · File: Exceptions.xlsx · Table: ExceptionRegister
  └─ Condition: length(outputs('List_rows_present_in_a_table')?['body/value'])  is greater than  0
       ├─ True
       │    ├─ Data Operation: Create HTML table
       │    │     From: value (List rows) · Columns: Custom
       │    │     PaymentID | StudentName | AmountPaid | Exception
       │    └─ Office 365 Outlook: Send an email (V2)
       │          To: reviewer address · Subject: Payment exceptions for review
       │          Body: one-line intro + Output of Create HTML table
       └─ False: (nothing)
```

## Steps

1. **Prepare the file.** Use `data/Exceptions.xlsx` from this repo (or export the exception register from Power BI: ⋯ → Export data → Excel, then Insert → Table and name it `ExceptionRegister`). Upload it to OneDrive as `Exceptions.xlsx`.
2. **Create → Scheduled cloud flow.** Name: *Weekly payment exception alert*. Repeat every 1 week on Monday at 9:00.
3. Add **List rows present in a table** with the settings above.
4. Add a **Condition** with the `length(...)` expression on the left, *is greater than*, `0` on the right.
5. Under **True**, add **Create HTML table** (Custom columns: PaymentID, StudentName, AmountPaid, Exception).
6. Add **Send an email (V2)** with the HTML table output in the body.
7. **Save → Test → Manually.** The email should contain **44 rows**.

## Evidence for the README

Save screenshots to `docs/screenshots/`:

- `flow-designer.png` – the whole flow in the designer
- `flow-email.png` – the email that arrived

## What would change for production

- Read directly from a SharePoint list or the published Power BI semantic model instead of an exported file.
- Add an **Approvals** step that routes each exception to the staff member in `RecordedBy`, and log the outcome.
- Note: *List rows present in a table* returns 256 rows by default; turn on pagination in Settings if the register grows.
