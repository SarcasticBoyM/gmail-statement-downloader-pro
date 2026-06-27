# Contributing

Thanks for improving Gmail Statement Downloader Pro.

## How to contribute

1. Fork the repository.
2. Create a new branch.
3. Make changes.
4. Test in Google Apps Script.
5. Open a pull request.

## Code style

- Keep Apps Script functions small and readable.
- Use clear logs for failures.
- Do not hardcode personal email addresses, account numbers, or passwords.
- Keep default bank queries generic.

## Testing checklist

- `setupBankConfig()` creates all sheets.
- `downloadCreditCards()` downloads only credit card statement PDFs.
- `downloadBankStatements()` downloads only bank statement PDFs.
- Duplicate files are skipped.
- `createCAPackage()` creates ZIP and CSV summary.
