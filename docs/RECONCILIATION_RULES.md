# Reconciliation Rules

## Normal Payments

- Search the bank account indicated by Payment Type.
- Match exact amount within ±3 calendar days.
- Prefer the same date, then the nearest eligible date.
- When duplicate exact-amount entries exist, use party-name vs bank-narration similarity.

## GPay

- Search Axis Bank 1213.
- Try individual exact-amount matches on the same date first.
- Then try individual exact-amount matches on the next date.
- Group only the remaining unmatched GPay entries.
- Never consume one bank transaction more than once.

## Cheques

- Extract the cheque number from Payment Type, for example `Cheque (000692)`.
- Normalize leading zeroes so `000692` and `692` are equivalent.
- Start with same-date exact amount.
- If the same date has duplicate exact-amount credits, use cheque number + amount.
- Search up to ±10 days when same-date matching fails.
- Use the extended cheque-number + amount fallback up to 30 days.
- Treat deposit → debit reversal → redeposit → final successful credit as one cheque lifecycle.
- Keep bounced presentations in the bounced-cheque report while using the final successful credit for reconciliation.

## Self Account

Recognized self/own-account transfers are excluded from unmatched entries and written to `SELF_ACCOUNT_ENTRIES`.

## Missing Vyapar Entries

The export sheet includes eligible unmatched bank credits and debits:

- Credit → Payment In
- Debit → Payment Out

Self-account transfers and cheque-bounce reversals are excluded.
