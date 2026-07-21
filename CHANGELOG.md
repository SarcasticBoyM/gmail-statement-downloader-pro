# Changelog

## 1.1.1 - 2026-07-21

### Fixed

- `Bank Cheque No`, `Cheque No`, and `Cheque/Reference No` report columns are now forced to plain-text format
- Removed accidental rupee/currency display from cheque-number columns
- Preserves leading zeroes in cheque numbers, such as `000568` and `000692`

## 1.1.0 - 2026-07-21

### Changed

- Updated `VYAPAR_MISSING_ENTRIES` to the requested import-ready layout
- Removed the `Payment Type` column from the import-ready report
- Added `Bank Account Type` with the exact source bank sheet (`1213` or `2224`)
- Report columns are now: `Vyapar Date`, `Type`, `Suggested Party Name`, `Amount`, `Bank Account Type`, `Bank Narration`
- Payment-In and Payment-Out remain included; self transfers and cheque-bounce reversals remain excluded

## 1.0.0 - 2026-07-21

Initial GitHub release based on script V17.

### Included

- Two-bank reconciliation
- Normal transaction ±3 day matching
- GPay individual-first and grouped-fallback matching
- Cheque number normalization
- Cheque deposit, bounce, redeposit, and final-clearance lifecycle
- Party/narration similarity matching
- Self-account separation
- Cheque-bounce report
- Party ledger summary with bank totals
- Export-ready missing Payment In and Payment Out sheet
