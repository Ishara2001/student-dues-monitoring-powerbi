# Copilot vs Claude – output evaluation

Give both tools the same file and the same prompt, then check every claim against the numbers already verified in the dashboard.

1. Upload `data/Exceptions.xlsx` to **Microsoft Copilot** (copilot.microsoft.com or Copilot in Excel) and to **Claude** (claude.ai).
2. Use exactly this prompt in both:

   > This is a register of payment exceptions from a student fees system. Summarise the main risks, give the count and total value for each exception type, and suggest three follow-up actions for the finance office.

3. Score both answers in `Copilot_vs_Claude_Scorecard.xlsx` (yellow cells). The correct answers are pre-filled from the dashboard.
4. Write three sentences of findings on the second sheet: which tool was more accurate, the most serious error either made, and how to prompt differently next time.
5. Save both raw answers as `copilot-answer.md` and `claude-answer.md` in this folder.

## Ground truth

| Exception | Payments | Value (LKR) |
|---|---:|---:|
| Possible duplicate | 18 | 47,000 |
| Missing receipt | 10 | 17,000 |
| Large cash payment | 7 | 126,000 |
| Overpayment | 5 | 425,000 |
| Future-dated | 4 | 8,000 |
| **Total** | **44** | **623,000** |

Things a good answer should notice: the 18 duplicate rows are 9 pairs; the 5 overpayments look like an extra zero typed (e.g. 180,000 for an 18,000 hostel fee); one overpayment is also a large cash payment; all 5 overpayments were recorded by K. Perera, all 4 future-dated entries by R. Fernando, and 7 of the 10 missing receipts by N. Silva.

## Findings

_To be completed after the test._

| | Copilot | Claude |
|---|---|---|
| Score (max 18) | | |
| Most serious error | | |

**Lesson:** AI output has to be checked against a trusted source before an auditor relies on it.
