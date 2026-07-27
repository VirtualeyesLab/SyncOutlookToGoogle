# Outlook to Google Calendar Sync

A Python utility that performs a one-way sync from a corporate Exchange/Outlook calendar to a personal Google Calendar. It uses Power Automate to create one local JSON file per event change, which is then synced via the Google Calendar API.

---

# Setup

### 1. Power Automate
Create a flow that writes one JSON file per event change to a folder in OneDrive.

- **Trigger:** `When an event is added, updated or deleted (V3)`
- **Create file:** filename should be unique per event (for example using event id).  
- **JSON content schema:**
  - `actionType` (`created`, `updated`, or `deleted`)  
  - `eventId`  
  - `subject`  
  - `start`  
  - `end`  
  - `location`  
  - `isAllDay`

Processed files are deleted only after successful sync. Failed files remain in the folder for retry.

---

### 2. Google API
- Create a project in Google Cloud Console  
- Enable **Google Calendar API**  
- Create OAuth client (Desktop app)  
- Download and save as `credentials.json` in the repository root  

---

### 3. Install
    pip install -r requirements.txt

---

## Usage

### GUI (recommended)
    python gui_app.py

- Select the Outlook event folder  
- Configure settings  
- Authenticate with Google  

### Agent (auto sync)
    python agent.py --monitor

### One-off sync
    python sync.py

### Purge previously synced Google events
Use this when you need to remove events created by this tool (for example after a timezone mapping fix) before a clean re-sync.

Preview only (no deletion):
    python agent.py --purge-synced --purge-dry-run

Delete tagged future events:
    python agent.py --purge-synced

---

## Auto Startup (Windows)

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

## File Monitoring

The agent uses two methods to detect changes in the event folder:

1. **Watchdog (Preferred):** Real-time file system notifications (efficient and responsive)
2. **Polling (Fallback):** Checks for new/changed `.json` files periodically (if watchdog unavailable)

The agent automatically selects the best available method and switches between them if needed.

---

## Troubleshooting

### "Snapshot file not found"
- Verify the path in config.json is correct
- Ensure your Power Automate flow is running and writing event JSON files
- Check that the folder exists at the configured location

### Authentication errors
- Delete `token.json` and re-authenticate through the GUI
- Verify your Google Cloud credentials.json is in the repository root

### No events being synced
- Check `sync.log` for detailed error messages
- Verify event JSON files are present in the configured folder
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