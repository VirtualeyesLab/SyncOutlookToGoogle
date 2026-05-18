# Outlook to Google Calendar Sync

A lightweight, stateless Python utility that performs a one-way sync from a corporate Exchange/Outlook calendar to a personal Google Calendar. 

This tool is specifically designed to bypass organizational policies that prevent external calendar sharing. It uses Power Automate to create a local snapshot of your Exchange calendar, which is then parsed and synced locally via the Google Calendar API.

## 🚀 Features
* **Stateless Synchronization:** Uses Google Calendar's invisible "Extended Properties" to track events. No local database required.
* **Duplicate Prevention:** Safely handles event creation, updates, and cancellations without duplicating entries.
* **Timezone Aware:** Correctly maps local Exchange times to Google Calendar timezones.
* **Secure:** All API communication happens locally. No cloud-to-cloud connections required.

---

## 🛠️ Phase 1: Power Automate Setup (The Exporter)
You need to create a flow in Power Automate to dump your upcoming events into a JSON file on your Business OneDrive.

1. **Trigger:** `When an event is added, updated or deleted (V3)` (or a Recurrence schedule).
2. **Action 1 - Get Events:** Add `Get calendar view of events (V3)`. 
   * **Start Time:** Use the expression `utcNow()`
   * **End Time:** Use the expression `addDays(utcNow(), 30)` (or however far out you want to sync).
3. **Action 2 - Select:** Add a Data Operations `Select` action. 
   * **From:** `value` (Dynamic Content)
   * **Map:** Create the following exact key-value pairs (Left side = Text, Right side = Dynamic Content):
     * `ExchangeID` ➡️ **Id**
     * `Subject` ➡️ **Subject**
     * `StartTime` ➡️ **Start time**
     * `EndTime` ➡️ **End time**
     * `Location` ➡️ **Location**
     * `IsAllDay` ➡️ **Is all day event**
     * `Categories` ➡️ **Categories**
4. **Action 3 - Create File:** Add a OneDrive for Business `Create file` action.
   * Save the output of the "Select" step (use expression `body('Select')` if it's hidden) to a file named `OutlookSnapshot.json`. Ensure it overwrites on each run.

---

## 🔐 Phase 2: Google API Setup (The Importer)
You need to authorize the script to edit your personal Google Calendar.

1. Go to the [Google Cloud Console](https://console.cloud.google.com/).
2. Create a new Project and enable the **Google Calendar API**.
3. Navigate to **Credentials** -> **Create Credentials** -> **OAuth client ID**.
4. Set Application Type to **Desktop app**.
5. Download the resulting JSON file, rename it to `credentials.json`, and place it in this repository's root folder.

---

## 💻 Phase 3: Local Environment Setup

1. **Clone the repository.**
2. **Install requirements:**
   ```bash
   pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib
   ```
3. **Update your `.gitignore`:**
   Before committing anything, ensure your existing `.gitignore` includes the following to protect your personal data and tokens:
   ```text
   # Google Sync Secrets
   credentials.json
   token.json
   config.json
   sync.log
   ```

---

## ▶️ Running the Sync

Run the script from your terminal:
```bash
python sync.py
```

### First-Run Setup:
1. The script will prompt you for the local path to your `OutlookSnapshot.json` file (e.g., `C:\Users\Name\OneDrive\CalendarSync\OutlookSnapshot.json`). It saves this to `config.json`.
2. A browser window will open asking you to authenticate with your Google account. Accept the permissions. This generates a local `token.json` file.

### Standard Execution:
Subsequent runs will happen silently in the background, outputting a high-level summary to the terminal and saving detailed logs to `sync.log`.

### Verbose Mode:
If you need to troubleshoot or see exactly which events are being updated/deleted, run the script with the verbose flag:
```bash
python sync.py -v
```
*(or `python sync.py --verbose`)*

---

## 🤖 Automation
To keep your calendars synced automatically, point **Windows Task Scheduler** to your python executable and pass this script as an argument to run every 15-30 minutes.