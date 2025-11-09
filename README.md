# Phil's Super Syncer

A Windows system tray application that automatically syncs Outlook calendar events to Google Calendar using Power Automate and Excel as an intermediary.

![.NET Framework 4.7.2](https://img.shields.io/badge/.NET%20Framework-4.7.2-blue)
![License](https://img.shields.io/badge/license-MIT-green)

## 📋 Overview

Phil's Super Syncer monitors an Excel file populated by Power Automate (tracking Outlook calendar changes) and automatically syncs those changes to your Google Calendar. The app runs silently in your system tray and provides a simple settings interface.

## ✨ Features

- 🔄 **Automatic Sync** - Monitors Excel file at configurable intervals (1-120 minutes)
- 🔐 **Secure OAuth** - Uses Google OAuth 2.0 for authentication
- 📊 **Excel Integration** - Reads event data from Power Automate-generated Excel files
- 🎯 **Event Tracking** - Maintains local mapping of Outlook → Google event IDs
- 🖥️ **System Tray App** - Runs in background with minimal UI
- ⚡ **Manual Sync** - Trigger sync on-demand from the tray menu
- 📝 **Smart Sync** - Handles added, updated, and deleted events
- 🕐 **All-Day Events** - Properly handles both timed and all-day events
- 📋 **Smart Logging** - Auto-enabled during setup, auto-disabled after first sync (configurable)
- 📁 **Log Access** - Easy access to logs from system tray for troubleshooting

## 📦 Requirements

### System Requirements
- **Windows 7 or later**
- **.NET Framework 4.7.2** or higher ([Download here](https://dotnet.microsoft.com/download/dotnet-framework/net472))

### Prerequisites
1. **Microsoft Power Automate** - Set up a flow to export Outlook calendar events to Excel
2. **Google Account** - With access to Google Calendar
3. **Google Cloud Console Access** - To create OAuth credentials
4. **OneDrive for Business** (or accessible file location) - For storing the Excel sync file

## 🚀 Installation

### Step 1: Clone or Download the Repository

```bash
git clone https://github.com/yourusername/SyncOutlookToGoogle.git
cd SyncOutlookToGoogle
```

### Step 2: Install NuGet Packages

Open the solution in **Visual Studio 2019** or later and restore NuGet packages. The project uses the following packages:

| Package | Version | Purpose |
|---------|---------|---------|
| `ClosedXML` | 0.104.2 | Excel file reading/writing |
| `Google.Apis.Calendar.v3` | Latest | Google Calendar API |
| `Google.Apis.Auth` | Latest | Google OAuth authentication |
| `Google.Apis.Oauth2.v2` | Latest | Google user info API |
| `Newtonsoft.Json` | 13.0.3 | JSON serialization |

**To restore packages in Visual Studio:**
1. Right-click on the solution in Solution Explorer
2. Select **"Restore NuGet Packages"**

**Or via Package Manager Console:**
```powershell
Install-Package ClosedXML -Version 0.104.2
Install-Package Google.Apis.Calendar.v3
Install-Package Google.Apis.Auth
Install-Package Google.Apis.Oauth2.v2
Install-Package Newtonsoft.Json -Version 13.0.3
```

### Step 3: Set Up Google OAuth Credentials

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project (or select an existing one)
3. Enable the following APIs:
   - **Google Calendar API**
   - **Google OAuth2 API**
4. Create OAuth 2.0 credentials:
   - Go to **APIs & Services** → **Credentials**
   - Click **"Create Credentials"** → **"OAuth client ID"**
   - Choose **"Desktop app"**
   - Download the JSON file
5. Rename the downloaded file to `client_secret.json`
6. Place `client_secret.json` in the **project root** (next to the `.exe` after building)

> ⚠️ **IMPORTANT**: Never commit `client_secret.json` to version control!

### Step 4: Set Up Excel Template

1. **Copy the Sample File**:
   - Locate `Sample Excel.xlsx` in the repository
   - Copy it to your **OneDrive for Business** folder (or another location accessible to Power Automate)
   - Rename it if desired (e.g., `Outlook-to-Google-Sync.xlsx`)

2. **Why OneDrive for Business?**
   - Power Automate can easily read/write to OneDrive for Business
   - File is automatically synced and backed up
   - Can be accessed from anywhere

> 💡 **Tip**: You can also use SharePoint, Teams, or any cloud storage location that Power Automate can access.

### Step 5: Build the Application

In Visual Studio:
1. Set build configuration to **Release**
2. Build → **Build Solution** (or press `Ctrl+Shift+B`)
3. The executable will be in `bin\Release\SyncOutlookToGoogle.exe`

### Download and Install

1. **Download the latest release** from [Releases](https://github.com/yourusername/SyncOutlookToGoogle/releases)
2. **Extract the ZIP file** to a folder (e.g., `C:\Program Files\PhilsSuperSyncer`)
3. **Windows SmartScreen Warning** - You may see a warning because this is a free, community app:
   
   **If you see "Windows protected your PC":**
   - Click **"More info"**
   - Click **"Run anyway"**
 
   This is normal for unsigned free software. The app is open-source and safe to run.

4. **First-time setup**: Follow the setup wizard to configure Excel and Google Calendar

> 💡 **Why does Windows warn about this app?**  
> Code signing certificates cost $100-400/year. Since this is a free app, it's not signed. Windows SmartScreen will learn to trust it as more people use it safely.

## 📊 Power Automate Setup

### Option 1: Use the Provided Excel Template

The repository includes **`Sample Excel.xlsx`** with the correct structure pre-configured:

1. **Upload to OneDrive for Business**:
   - Copy `Sample Excel.xlsx` to your OneDrive for Business folder
   - Note the file path/location

2. **Create Your Power Automate Flow**:
   - Create a new automated cloud flow
   - **Trigger**: "When an event is created or modified (V3)" (Outlook Calendar)
   - **Actions**:
     - Get event details from Outlook
     - Add a row to the Excel table (use your uploaded file)
     - Map Outlook fields to Excel columns (see table below)

3. **Important Settings**:
   - Excel table name must be **`Table1`** (already set in the sample)
   - Set `IsProcessed` to **`FALSE`** for all new rows
   - Use ISO 8601 format for all date/time fields

### Option 2: Create Your Own Excel File

If you prefer to create your own Excel file, it must contain a table named **`Table1`** with these columns:

| Column Name | Type | Description | Example |
|-------------|------|-------------|---------|
| `IsProcessed` | Text | "FALSE" for new events | `FALSE` |
| `Outlook Event ID` | Text | Unique Outlook event ID | `AAMkAGI1...` |
| `ActionType` | Text | "added", "updated", or "deleted" | `added` |
| `Timestamp` | Text | ISO 8601 timestamp | `2025-01-06T10:30:00Z` |
| `Subject` | Text | Event title | `Team Meeting` |
| `Body` | Text | Event description | `Quarterly review` |
| `Location` | Text | Event location | `Conference Room A` |
| `StartTime` | Text | ISO 8601 start time | `2025-01-06T14:00:00Z` |
| `EndTime` | Text | ISO 8601 end time | `2025-01-06T15:00:00Z` |
| `isAllDay` | Text | "true" or "false" | `false` |

**To create the table in Excel:**
1. Open a new Excel workbook
2. Add the column headers in row 1
3. Select the header row (and a few empty rows below)
4. Go to **Insert** → **Table**
5. Check **"My table has headers"**
6. Name the table **`Table1`** (in Table Design tab)
7. Save the file to OneDrive for Business

**Example Excel row:**
```
IsProcessed: FALSE
Outlook Event ID: AAMkAGI1AAAAAA=
ActionType: added
Timestamp: 2025-01-06T10:30:00Z
Subject: Doctor Appointment
Body: Annual checkup
Location: Medical Center
StartTime: 2025-01-10T09:00:00-05:00
EndTime: 2025-01-10T10:00:00-05:00
isAllDay: false
```

### Power Automate Field Mapping

When setting up your Power Automate flow, map the Outlook fields to Excel columns:

| Excel Column | Power Automate Expression/Field |
|--------------|--------------------------------|
| `IsProcessed` | `FALSE` (hardcoded) |
| `Outlook Event ID` | `triggerOutputs()?['body/id']` |
| `ActionType` | `added` (or `updated`, `deleted` based on trigger) |
| `Timestamp` | `utcNow()` |
| `Subject` | `triggerOutputs()?['body/subject']` |
| `Body` | `triggerOutputs()?['body/body/content']` |
| `Location` | `triggerOutputs()?['body/location/displayName']` |
| `StartTime` | `triggerOutputs()?['body/start/dateTime']` + `Z` |
| `EndTime` | `triggerOutputs()?['body/end/dateTime']` + `Z` |
| `isAllDay` | `triggerOutputs()?['body/isAllDay']` |

## 🎮 Usage

### First-Time Setup

1. **Run the application** - The settings window will appear
2. **Select Excel File**:
   - Click **"Change..."** in section 1
   - Browse to your OneDrive for Business folder
   - Select your sync Excel file (e.g., `Outlook-to-Google-Sync.xlsx`)
 - A green ✔ should appear if the file is valid
3. **Login to Google**:
   - Click **"Login to Google"** in section 2
   - A browser window will open
   - Sign in and authorize the application
   - Select your target Google Calendar from the dropdown
4. **Configure Sync Interval**:
   - Set how often to check for changes (default: 5 minutes)
5. **Click "Save"**:
   - Settings are saved
   - App minimizes to system tray
   - Automatic syncing begins

### Daily Use

Once configured, the app runs automatically in the system tray:

- **System Tray Icon**: Right-click for options
  - **Sync Now** - Manually trigger a sync
  - **Show Settings** - Open the settings window
  - **Open Logs Folder** - View application logs for troubleshooting
  - **Exit** - Close the application

- **Settings Window**: Double-click the tray icon to open
  - View sync status
  - Change sync interval
  - **Enable/disable detailed logging** (new checkbox in Section 3)
  - Logout/Reset

### Sync Behavior

- The app checks the Excel file every `X` minutes (configurable)
- Only rows with `IsProcessed = FALSE` are synced
- After successful sync, rows are marked `IsProcessed = TRUE`
- Event mappings are stored in `id_map.json` for tracking updates/deletes

## 📁 Files Created by the App

The application creates these files in its directory:

| File | Purpose | In .gitignore? |
|------|---------|----------------|
| `settings.json` | Stores your configuration (refresh token, file paths, etc.) | ✅ Yes |
| `id_map.json` | Maps Outlook Event IDs → Google Event IDs | ✅ Yes |
| `token.json/` | Google OAuth token cache | ✅ Yes |
| `Logs/` | Daily log files (auto-disabled after first sync) | ✅ Yes |

> ⚠️ These files contain sensitive data and should **never** be committed to version control!

### Logging Behavior

- **During Setup**: Logging is automatically enabled to capture setup details
- **After First Sync**: Logging automatically disables to save disk space
- **Manual Control**: Re-enable anytime via "Enable detailed logging" checkbox in Settings
- **Log Location**: `Logs/` folder next to the executable
- **Quick Access**: Right-click tray icon → "Open Logs Folder"

## 🔧 Troubleshooting

### Issue: "Failed to authenticate with Google"
**Solution**: 
- Click **"Reset All"** button
- Delete the `token.json` folder manually
- Re-login to Google

### Issue: Excel validation shows red ✘
**Solution**:
- Ensure the Excel file exists
- Verify it contains a table named `Table1`
- Check that all required column headers are present (case-insensitive)
- Close Excel if the file is open
- If using OneDrive, ensure the file is fully synced locally

### Issue: Events not syncing
**Solution**:
- **Open Logs**: Right-click tray icon → "Open Logs Folder" to view detailed logs
- Enable logging in Settings if it's currently disabled
- Verify `IsProcessed` column has `FALSE` (not blank)
- Ensure timestamps are valid ISO 8601 format
- Check that the sync interval has elapsed
- Verify Power Automate is successfully writing to the Excel file
- Review log file for specific error messages

### Issue: Excel file is locked or "in use"
**Solution**:
- Close Excel if you have the file open
- OneDrive may be syncing - wait a few seconds and try again
- Check if another process has the file open (Task Manager → Performance → Open Resource Monitor)

### Issue: Icon not showing in system tray
**Solution**:
- Ensure `icon.ico` is embedded as a resource in the project
- Check **Build Action** is set to **"Embedded Resource"**
- Rebuild the project

### Issue: OneDrive sync conflicts
**Solution**:
- The app reads the file into memory, so brief lock times are minimal
- If conflicts occur, increase the sync interval to give OneDrive time to sync
- Consider using a local folder that OneDrive syncs, rather than the cloud-only location

## 🛠️ Development

### Project Structure

```
SyncOutlookToGoogle/
├── SyncOutlookToGoogle/
│   ├── Program.cs          # Main form and sync logic
│   ├── GoogleAuth.cs       # Google OAuth helper
│   ├── Logger.cs       # File-based logging system
│   ├── icon.ico            # Application icon (embedded)
│   ├── App.config          # Application configuration
│   └── Properties/
│       ├── AssemblyInfo.cs
│       └── Resources.resx
├── Sample Excel.xlsx       # Template Excel file with correct structure
├── .gitignore              # Excludes sensitive files
└── README.md    # This file
```

### Building from Source

1. Open `SyncOutlookToGoogle.sln` in Visual Studio
2. Restore NuGet packages (automatic)
3. Place `client_secret.json` in the output directory
4. Press **F5** to run in debug mode

### Debug Output

The application writes detailed logs to timestamped log files:
- **Location**: `Logs/` folder in the application directory
- **Format**: `PhilsSuperSyncer_YYYY-MM-DD.log`
- **Access**: Right-click tray icon → "Open Logs Folder"
- **Automatic Cleanup**: Keeps last 10 log files, rotates when exceeding 5MB
- **Smart Logging**: Auto-enabled during setup, auto-disabled after first sync

**What's Logged:**
- Sync operations
- Excel file processing
- Google API calls
- Error messages
- Settings changes

**Enabling Logs:**
- Open Settings → Section 3: Sync Settings
- Check "Enable detailed logging"
- Click "Save"

Use logs when troubleshooting issues or reporting bugs.

## 🔒 Security & Privacy

- **OAuth tokens** are stored locally in `settings.json` and `token.json/`
- **No data** is sent to any third-party services (only Google Calendar API)
- **Excel file** remains local and is never uploaded
- **ID mappings** are stored locally in `id_map.json`

### Important Security Notes

1. **Never commit** `client_secret.json` to version control
2. **Keep** `settings.json` and `id_map.json` private
3. **Revoke access** in [Google Account Settings](https://myaccount.google.com/permissions) if needed
4. **Reset credentials** if you suspect they've been compromised

## 📜 License

This project is licensed under the MIT License - see below for details:

```
MIT License

Copyright (c) 2025 Dr Phil Turnbull

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
```

## 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

### Ideas for Contributions
- [ ] Add retry logic for failed API calls
- [ ] Support multiple Google calendars
- [x] Add logging to file instead of Debug output ✅ **DONE!**
- [ ] Create installer (e.g., with WiX)
- [ ] Add support for recurring events
- [ ] Implement two-way sync (Google → Outlook)
- [ ] Add email notifications for sync failures
- [ ] Support for event reminders/notifications

## 🙏 Acknowledgments

- **ClosedXML** - For excellent Excel file handling
- **Google APIs** - For Calendar and OAuth support
- **Newtonsoft.Json** - For JSON serialization

## 📞 Support

If you encounter issues:
1. Check the **Troubleshooting** section above
2. Review **Debug output** for error messages
3. Open an [Issue](https://github.com/yourusername/SyncOutlookToGoogle/issues) on GitHub

---

**Made with ❤️ for anyone tired of manual calendar syncing!**
