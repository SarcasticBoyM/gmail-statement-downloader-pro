# GitHub par kaise upload karein

## Web browser se

1. GitHub par login karein.
2. New repository banayein: `gmail-statement-downloader-pro`
3. Public ya Private select karein.
4. Repository create karein.
5. `Add file` → `Upload files` par click karein.
6. Is ZIP ko extract karke saare files drag and drop karein.
7. `Commit changes` click karein.

## Command line se

```bash
git init
git add .
git commit -m "Initial release v2.1.0"
git branch -M main
git remote add origin https://github.com/USERNAME/gmail-statement-downloader-pro.git
git push -u origin main
```
