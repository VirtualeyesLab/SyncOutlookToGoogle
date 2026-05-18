# Outlook to Google Calendar Sync

A lightweight, stateless Python utility that performs a one-way sync from a corporate Exchange/Outlook calendar to a personal Google Calendar. It uses Power Automate to create a local snapshot of your Exchange calendar, which is then parsed and synced locally via the Google Calendar API.

## Features
* **Stateless Synchronization:** Uses Google Calendar's invisible "Extended Properties" to track events. No local database required.
* **Duplicate Prevention:** Safely handles event creation, updates, and cancellations without duplicating entries.
* **Timezone Aware:** Correctly maps local Exchange times to Google Calendar timezones.
* **Secure:** All API communication happens locally. No cloud-to-cloud connections required.

---

## Setup
Not there is a bit of initial configuration to bridge the Exchange space to the Google space. The intermediary file is human readable JSON, though should not be manually edited. 

## 1: Power Automate Setup 
You need to create a flow in Power Automate to dump your upcoming events into a JSON file on your Business OneDrive.

1. **Trigger:** `When an event is added, updated or deleted (V3)` .
2. **Action 1 - Get Events:** Add `Get calendar view of events (V3)`. 
   * **Start Time:** Use the expression `utcNow()`
   * **End Time:** Use the expression `addDays(utcNow(), 360)` (or however far out you want to sync).
3. **Action 2 - Select:** Add a Data Operations `Select` action. 
   * **From:** `value` (Dynamic Content)
   * **Map:** Create the following exact key-value pairs (Left side = Text, Right side = Dynamic Content):
     * `ExchangeID` - **Id**
     * `Subject` - **Subject**
     * `StartTime` - **Start time**
     * `EndTime` - **End time**
     * `Location` - **Location**
     * `IsAllDay` - **Is all day event**
     * `Categories` - **Categories**
4. **Action 3 - Create File:** Add a OneDrive for Business `Create file` action.
   * Save the output of the "Select" step (use expression `body('Select')` if it's hidden) to a file named `OutlookSnapshot.json`. Ensure it overwrites on each run.

---

## 2: Google API Setup 
You need to authorise the script to edit your personal Google Calendar. This part is a bit tricky, but there's a Cloud Gemini Assistant who can help. Note that you must download the JSON file IMMEDIATELY - the option disappears otherwise. If this happens, no big deal, just generate a new Secret and download that JSON. 

1. Go to the [Google Cloud Console](https://console.cloud.google.com/).
2. Create a new Project and enable the **Google Calendar API**.
3. Navigate to **Credentials** -> **Create Credentials** -> **OAuth client ID**.
4. Set Application Type to **Desktop app**.
5. Download the resulting JSON file IMMEDIATELY, rename it to `credentials.json`, and place it in this repository's root folder.

---

## 3: Local Environment Setup

1. **Clone the repository.**
2. **Install all requirements:**
   ```bash
   pip install -r requirements.txt
   ```
   This includes:
   - Google Calendar API client
   - watchdog (for file monitoring)
   - PyQt5 (for GUI)

---

## 4: Running the Sync

### Option A: GUI Settings & Monitoring (Recommended)

The easiest way to manage the sync is through the graphical interface:

1. **Launch the GUI:**
   ```bash
   python gui_app.py
   ```

2. **Configure settings:**
   - Navigate to the "Settings" tab
   - Click "Browse..." to select your `OutlookSnapshot.json` file
   - Adjust timezone, sync frequency, and logging level as needed
   - Click "Save Settings"

3. **Authenticate with Google:**
   - Go to the "Google Authentication" tab
   - Click "Authenticate"
   - A browser window will open; authorize the app to access your Google Calendar
   - Return to the app to confirm successful authentication

4. **View sync history:**
   - The "Sync History" tab shows the last sync results
   - View detailed logs to troubleshoot any issues

### Option B: Command-Line Monitoring Agent

Run the monitoring agent that automatically triggers syncs when the snapshot file changes:

```bash
python agent.py --monitor
```

The agent will:
- Monitor your `OutlookSnapshot.json` file for changes (using watchdog)
- Automatically trigger a sync when changes are detected
- Fall back to polling if watchdog is unavailable
- Log all activity to `sync.log`

### Option C: Manual Single Sync

Run a single sync immediately:

```bash
python sync.py
```

Or via the agent:
```bash
python agent.py --sync-now
```

---

## 5: Automatic Startup

To run the agent automatically when Windows starts:

### Using the Setup Script (Recommended)

Run as Administrator:
```bash
python setup_scheduler.py --install
```

This creates a Windows Task Scheduler task that runs the agent at user logon.

**To uninstall:**
```bash
python setup_scheduler.py --uninstall
```

**To check status:**
```bash
python setup_scheduler.py --status
```

---

## Configuration

Settings are stored in `config.json`:

```json
{
    "OUTLOOK_JSON_PATH": "C:\\Users\\Name\\OneDrive\\OutlookSnapshot.json",
    "TIMEZONE": "Pacific/Auckland",
    "SYNC_FREQUENCY_MINUTES": 15,
    "MONITORING_ENABLED": true,
    "LOGGING_LEVEL": "INFO",
    "LAST_SYNC_TIME": "2024-01-15T14:30:45.123456",
    "LAST_SYNC_STATUS": "Success",
    "LAST_SYNC_CREATED": 5,
    "LAST_SYNC_UPDATED": 3,
    "LAST_SYNC_DELETED": 1
}
```

You can edit this file directly, but the GUI is recommended.

---

## File Monitoring

The agent uses two methods to detect changes to `OutlookSnapshot.json`:

1. **Watchdog (Preferred):** Real-time file system notifications (efficient and responsive)
2. **Polling (Fallback):** Checks file modification time periodically (if watchdog unavailable)

The agent automatically selects the best available method and switches between them if needed.

---

## Troubleshooting

### "Snapshot file not found"
- Verify the path in config.json is correct
- Ensure your Power Automate flow is running and updating `OutlookSnapshot.json`
- Check that the file exists at the configured location

### Authentication errors
- Delete `token.json` and re-authenticate through the GUI
- Verify your Google Cloud credentials.json is in the repository root

### No events being synced
- Check `sync.log` for detailed error messages
- Verify your Outlook events exist in the snapshot file
- Ensure the configured timezone matches your local timezone

### GUI won't start
- Verify PyQt5 is installed: `pip install PyQt5>=5.15`
- Try running from command line to see detailed error messages

---

## Files

| File | Purpose |
|------|---------|
| `sync.py` | Core sync logic (can be run standalone) |
| `agent.py` | Background agent with file monitoring |
| `gui_app.py` | PyQt5 graphical interface |
| `agent_config.py` | Configuration manager and schema |
| `setup_scheduler.py` | Windows Task Scheduler integration |
| `config.json` | Settings and sync history |
| `token.json` | Google authentication token (gitignored) |
| `credentials.json` | Google OAuth credentials (gitignored) |
| `sync.log` | Detailed execution logs |

---

## Architecture

The sync system consists of three components:

- **Core Sync Engine (`sync.py`):** Handles Outlook→Google synchronization logic
- **Monitoring Agent (`agent.py`):** Watches for file changes and triggers syncs
- **GUI Manager (`gui_app.py`):** Provides user-friendly configuration and monitoring
- **Config Manager (`agent_config.py`):** Persists and manages all settings

All components share the same configuration file (`config.json`) for centralized settings management.

---

##  Development

The refactored architecture allows easy extension:

- Extracted sync functions can be imported and used programmatically
- ConfigManager provides centralized settings
- SyncAgent can be used as a library in other Python apps
- GUI is independent and can be run separately

Example programmatic usage:
```python
from sync import perform_sync, get_google_service, load_outlook_snapshot
from agent_config import ConfigManager

config = ConfigManager()
service = get_google_service(config.get_timezone())
events = load_outlook_snapshot(config.get_outlook_json_path())
stats = perform_sync(service, events, config.get_timezone())
```