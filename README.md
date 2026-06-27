# Gmail Statement Downloader Pro

A no-Python, Google Apps Script tool that downloads credit card statements and bank account statements from Gmail into Google Drive, bank-wise, with duplicate tracking, logs, dashboard, and a one-click CA package ZIP.

## Features

- Credit card statement download
- Bank account statement download
- Bank add, remove, edit from Google Sheet config
- Duplicate detection using Gmail message ID + attachment name
- Bank-wise Google Drive folders
- Logs, history, and dashboard sheets
- One Click CA Package ZIP
- CSV summary for CA/auditor
- No Python or desktop installation required

## Quick Start

1. Open Google Apps Script: https://script.google.com
2. Create a new project.
3. Copy `src/Code.gs` into the Apps Script editor.
4. Save the project.
5. Run `setupBankConfig` once.
6. Allow Gmail and Drive permissions.
7. Open the config sheet link from Execution log.
8. Edit banks in the `Banks` sheet.
9. Run `downloadAllStatements` or `createCAPackage`.

## Main Functions

| Function | Purpose |
|---|---|
| `setupBankConfig()` | Creates/opens config spreadsheet |
| `downloadCreditCards()` | Downloads only credit card statements |
| `downloadBankStatements()` | Downloads only bank account statements |
| `downloadAllStatements()` | Downloads all enabled statement types |
| `createCAPackage()` | Downloads missing statements and creates a CA ZIP |
| `openDownloadFolder()` | Logs the Drive folder link |
| `refreshDashboard()` | Updates dashboard stats |

## Bank Configuration

Open the generated Google Sheet and edit the `Banks` tab.

| Column | Example | Notes |
|---|---|---|
| Bank | HDFC Bank | Bank display name |
| Type | CREDIT_CARD | Use `CREDIT_CARD` or `BANK_STATEMENT` |
| Query | from:hdfcbank filename:pdf statement | Gmail search query |
| Enabled | YES | Set `NO` to disable |
| Notes | optional | Your notes |

## CA Package

Run `createCAPackage()` to create a ZIP file in Google Drive. The ZIP includes statement PDFs and `CA_Statement_Summary.csv`.

## Folder Structure in Drive

```text
Gmail Statement Downloader Pro/
├── Credit Cards/
├── Bank Statements/
└── CA Packages/
```

## Privacy & Security

This tool runs inside your own Google account using Google Apps Script. It does not send statements to any external server. Files are saved only in your Google Drive.

## Troubleshooting

### `Cannot call SpreadsheetApp.getUi()`
This happens in standalone script context. The project logs the sheet/folder link in Execution log.

### No statements downloaded
Check that:
- `Enabled` is `YES`
- Gmail query is correct
- matching emails have PDF attachments
- date filters in Settings are blank or valid

## Roadmap

- Sidebar UI
- Progress bar
- Scheduled monthly downloads
- One-click email to CA
- Excel summary
- More default bank query templates

## License

MIT
