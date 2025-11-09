# Pre-GitHub Push Checklist

## ? **Files to Review Before First Commit**

### **CRITICAL - Never Commit These:**
- [ ] `client_secret.json` - **DELETE if present in repo!**
- [ ] `settings.json` - Should be in .gitignore ?
- [ ] `id_map.json` - Should be in .gitignore ?
- [ ] `token.json/` folder - Should be in .gitignore ?
- [ ] `Logs/` folder - Should be in .gitignore ?
- [ ] Any `*.xlsx` files (user data) - Should be in .gitignore ?
- [ ] `bin/` and `obj/` folders - Should be in .gitignore ?

### **Files to Include:**
- [x] `README.md` - ? Updated with logging features
- [x] `.gitignore` - ? Comprehensive and up-to-date
- [x] `INSTALLATION_GUIDE.txt` - ? Already created
- [x] `LICENSE` or MIT License in README - ? In README
- [x] `Sample Excel.xlsx` - **MUST INCLUDE** (template for users)
- [x] Source code files (.cs) - ? All ready
- [x] Project files (.csproj, .sln) - ? Ready
- [x] `icon.ico` - ? Embedded resource

---

## ?? **Final Verification Steps**

### **1. Check Git Status**
```bash
cd C:\git\SyncOutlookToGoogle
git status
```

**Look for:**
- ? NO `client_secret.json`
- ? NO `settings.json`, `id_map.json`, `token.json`
- ? NO `*.log` files or `Logs/` folder
- ? NO `bin/` or `obj/` folders
- ? YES to all source files (.cs)
- ? YES to `Sample Excel.xlsx`

### **2. Test .gitignore**
```bash
# This should show ONLY files you want to commit
git add -A --dry-run
```

### **3. Create .gitattributes (Optional but Recommended)**
Create a file named `.gitattributes` in repo root:
```
# Auto-normalize line endings
* text=auto

# Explicit files
*.cs text
*.csproj text
*.sln text
*.md text
*.json text
*.txt text
*.config text

# Binary files
*.ico binary
*.xlsx binary
*.png binary
*.jpg binary
*.dll binary
*.exe binary
```

---

## ?? **Recommended First Commit Message**

```
Initial commit: Phil's Super Syncer v1.0

- Outlook to Google Calendar sync tool
- Uses Power Automate + Excel as intermediary
- System tray application with auto-sync
- Smart logging system (auto-disables after setup)
- OAuth 2.0 authentication
- ID mapping for event tracking
- Sample Excel template included

Built with .NET Framework 4.7.2
```

---

## ?? **GitHub First Push Commands**

```bash
# Initialize git (if not done yet)
git init

# Add all files (respecting .gitignore)
git add .

# Review what will be committed
git status

# Create first commit
git commit -m "Initial commit: Phil's Super Syncer v1.0"

# Create GitHub repo (do this on GitHub.com first!)
# Then connect your local repo:
git remote add origin https://github.com/yourusername/SyncOutlookToGoogle.git

# Push to GitHub
git branch -M main
git push -u origin main
```

---

## ?? **Post-Push Tasks**

### **On GitHub.com:**
1. **Add Topics/Tags:**
   - `calendar-sync`
   - `outlook`
   - `google-calendar`
   - `power-automate`
   - `csharp`
   - `dotnet-framework`
   - `windows`
   - `system-tray`

2. **Edit Description:**
   ```
   ?? Sync Outlook calendar to Google Calendar using Power Automate + Excel. Windows tray app with OAuth 2.0. One-way sync with smart logging.
   ```

3. **Add README Screenshot (later):**
   - Take screenshot of Settings window
   - Add to repo as `docs/screenshot.png`
   - Update README with `![Screenshot](docs/screenshot.png)`

4. **Create First Release (optional):**
   - Go to Releases ? "Create a new release"
   - Tag: `v1.0.0`
   - Title: "Phil's Super Syncer v1.0.0 - Initial Release"
   - Upload: ZIP file with compiled .exe + dependencies
   - Include: `INSTALLATION_GUIDE.txt`

---

## ?? **FINAL CHECK - Run These Commands**

```bash
# 1. Check for sensitive files
git ls-files | findstr /i "client_secret settings.json id_map token"
# Should return: NOTHING

# 2. Check if .gitignore is working
git check-ignore -v client_secret.json
# Should return: .gitignore:X:client_secret.json

# 3. Verify Sample Excel is included
git ls-files | findstr /i "Sample"
# Should return: Sample Excel.xlsx
```

---

## ? **All Clear for Push!**

If all checks pass:
```bash
git push -u origin main
```

**Your repository is ready for the world! ??**

---

## ?? **Security Reminder**

If you **accidentally commit** sensitive files:
1. **Immediately delete the repo** on GitHub
2. **Revoke Google OAuth credentials** in Google Cloud Console
3. **Generate new credentials**
4. **Start fresh**

Better safe than sorry! ???
