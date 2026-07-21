# Google Sheets Bank Reconciliation for Vyapar

A custom Google Apps Script that reconciles two Axis Bank statement sheets with Vyapar payment entries inside the same Google Sheets workbook.

The script is designed for business reconciliation workflows that need date-tolerance matching, cheque lifecycle handling, GPay grouping, party-name matching, bounced-cheque detection, self-account separation, and export-ready missing-entry reports.

## Features

- Reconciles two bank statement sheets against `VYAPAR_PAYMENT_IN`
- Normal transactions matched within **±3 days**
- Payment Type decides the target bank account
- GPay matching:
  - same-day individual match first
  - next-day individual match second
  - grouped match only for remaining entries
- Cheque matching:
  - same-date amount match first
  - duplicate amount resolved using cheque number
  - searches up to ±10 days
  - extended cheque-number + amount search up to 30 days
  - leading-zero-safe cheque numbers (`000692` = `692`)
  - supports deposit → bounce → redeposit → final clearance lifecycle
- Party name vs bank narration similarity for duplicate/tie resolution
- Detects and separates cheque-bounced transactions
- Detects and separates self-account transfers
- Generates party-wise Vyapar and bank summary
- Generates an import-ready missing-entry report for Payment-In and Payment-Out
- Prevents one bank transaction from being matched more than once
- Writes cheque/reference numbers as plain text, without currency formatting, while preserving leading zeroes

## Required Workbook Sheets

The Google Sheets workbook must contain these tabs:

1. `Balaji Traders Axis Bank-1213`
2. `Balaji Traders Axis Bank-2224`
3. `VYAPAR_PAYMENT_IN`

The workbook file name can be anything. Only the sheet/tab names above are used by default.

## Supported Vyapar Columns

The script supports the current column format:

- `Date`
- `Reference No`
- `Party Name`
- `Type`
- `Payment Type`
- `Received`

It also contains aliases for older headings such as `Ref No`, `Party`, `Entry Type`, and `Total Amt`.

## Generated Report Sheets

Running the reconciliation creates or refreshes:

- `RECONCILIATION_REPORT`
- `UNMATCHED_BANK_ENTRIES`
- `SELF_ACCOUNT_ENTRIES`
- `CHEQUE_BOUNCED_ENTRIES`
- `PARTY_LEDGER_SUMMARY`
- `VYAPAR_MISSING_ENTRIES`

The import-ready sheet contains:

- `Vyapar Date`
- `Type`
- `Suggested Party Name`
- `Amount`
- `Bank Account Type`
- `Bank Narration`

`Payment Type` is intentionally excluded. `Bank Account Type` shows the exact source sheet: `Balaji Traders Axis Bank-1213` or `Balaji Traders Axis Bank-2224`.

## Installation

1. Open the target Google Sheets workbook.
2. Go to **Extensions → Apps Script**.
3. Delete the existing content in `Code.gs`.
4. Copy the contents of this repository's `Code.gs` into the Apps Script editor.
5. Save the project.
6. Run `runReconciliation` once and allow the requested permissions.
7. Reload the Google Sheet.
8. Use **RECO → Run Reconciliation**.

## Configuration

Sheet names and matching thresholds are at the top of `Code.gs` inside `RECO_CONFIG`.

Common settings:

```javascript
const RECO_CONFIG = {
  SHEET_1213: 'Balaji Traders Axis Bank-1213',
  SHEET_2224: 'Balaji Traders Axis Bank-2224',
  SHEET_VYAPAR: 'VYAPAR_PAYMENT_IN',
  DIRECT_DATE_TOLERANCE_DAYS: 3,
  CHEQUE_PRIMARY_DAYS: 10,
  CHEQUE_FALLBACK_DAYS: 30,
  GPAY_MAX_DAY_OFFSET: 1
};
```

## Important Notes

- Test the script on a copy of the workbook before using it on live accounting data.
- Do not publish real bank statements, customer data, account numbers, or transaction exports in this repository.
- Keep sample files anonymized.
- The script does not send workbook data to an external API.

## Version

Current repository version: **1.1.0**, based on reconciliation script V18.

## Disclaimer

This script assists with reconciliation but does not replace accountant review. Review unmatched, manual-check, cheque-bounced, and export-ready entries before posting them into accounting software.
