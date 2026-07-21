# Upload This Project to GitHub

## Browser method

1. Create a new empty repository on GitHub.
2. Choose **Add file → Upload files**.
3. Drag all files and folders from this project directory into the upload area.
4. Enter commit message: `Initial release of bank reconciliation script`.
5. Commit directly to the `main` branch.

## Git command method

Run these commands inside the project folder:

```powershell
git init
git add .
git commit -m "Initial release of bank reconciliation script"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/google-sheets-bank-reconciliation.git
git push -u origin main
```

Create the empty GitHub repository before running the final two commands. Do not initialize the remote repository with a README if you are pushing this prepared folder, because this folder already contains a README.
