# Gmail Statement Downloader Pro - Hindi Guide

Ye tool bina Python install kiye Google Apps Script me chalega.

## Setup

1. https://script.google.com open karein.
2. New Project banayein.
3. `src/Code.gs` ka pura code copy karke Apps Script me paste karein.
4. Save karein.
5. Function dropdown me `setupBankConfig` select karke Run karein.
6. Gmail/Drive permission allow karein.
7. Execution log me Google Sheet link milega, use open karein.
8. `Banks` sheet me bank add/remove/edit karein.

## Use

- Sirf Credit Card: `downloadCreditCards`
- Sirf Bank Statement: `downloadBankStatements`
- Dono: `downloadAllStatements`
- CA ke liye ZIP: `createCAPackage`

## Bank Add / Remove

`Banks` sheet me:

- Naya bank add karna ho to new row add karein.
- Bank disable karna ho to `Enabled = NO` karein.
- Query improve karni ho to `Query` column edit karein.

## CA Package

`createCAPackage()` run karein. Google Drive me ZIP banegi jisme PDFs aur CSV summary hogi.
