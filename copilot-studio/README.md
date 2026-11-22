# Fee Policy Assistant (Copilot Studio)

An AI agent that answers questions about the fee policy using only `docs/Fee_Payment_Policy.pdf`, tested with 10 reusable test cases.

**Account:** Copilot Studio's free trial needs a work or school account. Sign up at [copilotstudio.microsoft.com](https://copilotstudio.microsoft.com) with the university account. The trial lets you build and test in the test chat panel; publishing is not needed. See [Sign up for a Copilot Studio trial](https://learn.microsoft.com/microsoft-copilot-studio/sign-up-individual).

## Setup

1. **Create → New agent**, name: **Fee Policy Assistant**.
2. **Instructions** (paste exactly):

   ```
   You answer questions from students and finance staff using only the Student Fee Payment Policy.
   Quote the section number. If the policy does not cover the question, say so and suggest
   contacting the finance office. Never guess.
   ```

3. **Knowledge** → upload `docs/Fee_Payment_Policy.pdf`. Turn **off** general/web knowledge so the agent answers only from the policy.

## Testing

Run the 10 questions in `Agent_Test_Cases.xlsx` in the test chat and record Pass / Partial / Fail under *Run 1*. For any Fail, adjust the instructions and run all 10 again under *Run 2*. Record the before and after scores, for example "7/10 → 10/10".

Questions 8 and 10 ask about a late fine and an early-payment discount, which the policy deliberately does not mention. They test whether the agent invents policy — the most important failure for a GRC assistant.

## Results

| Run | Score | What changed |
|---|---|---|
| 1 | _/10 | baseline instructions |
| 2 | _/10 | |

Screenshots: `docs/screenshots/agent-test-chat.png`.
