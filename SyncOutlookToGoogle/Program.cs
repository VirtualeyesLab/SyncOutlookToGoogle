using ClosedXML.Excel;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CalendarSyncEngine
{
    // ### SETTINGS HELPER CLASS ###
    public class SyncSettings
    {
        public string ExcelFilePath { get; set; }
        public string GoogleCalendarId { get; set; }
        public string GoogleCalendarName { get; set; }
        public string GoogleAccountEmail { get; set; }
        public string GoogleRefreshToken { get; set; }
        public int SyncIntervalMinutes { get; set; } = 5;
        public DateTime? LastSyncTime { get; set; }
        public bool EnableLogging { get; set; } = true; // Default: enabled
        public bool HasCompletedFirstSync { get; set; } = false; // Track first successful sync
    }

    // Helper class for our ComboBox
    public class CalendarItem
    {
        public string Id { get; set; }
        public string Summary { get; set; }
        public override string ToString() => Summary;
    }

    public partial class SyncEngineForm : Form
    {
        private NotifyIcon trayIcon;
        private ContextMenuStrip trayMenu;
        private System.Threading.Timer _syncTimer;

        // ### CONFIGURATION & STATE ###
        private static readonly object _syncLock = new object();
        private static bool _isSyncing = false;
        private const string SettingsFileName = "settings.json";

        // ### JSON DATABASE FILE FOR OUTLOOK GOOGLE EVENT LINKAGES ###
        private const string IdMapFileName = "id_map.json";
        private static Dictionary<string, string> _idMap = new Dictionary<string, String>();
        private static readonly object _idMapLock = new object(); // Lock for saving the map
        private static SyncSettings _settings;
        private const string ExcelTableName = "Table1";

        // ### UI CONTROLS ###
        private GroupBox grpExcel;
        private Label lblExcelFile;
        private Button btnChangeExcel;
        private Label lblExcelStatus;

        private GroupBox grpGoogle;
        private Label lblGoogleAccount;
        private Label lblGoogleCalendar;
        private Label lblCalendarName;
        private Button btnLoginGoogle;
        private Button btnLogoutGoogle;
        private ComboBox cmbGoogleCalendar; // Used only during login

        private GroupBox grpSync;
        private Label lblSyncInterval;
        private NumericUpDown numSyncInterval;
        private Button btnSyncNow;
        private Button btnSave;
        private Label lblLastSync;
        private CheckBox chkEnableLogging; // NEW: Logging checkbox

        private Button btnResetApp;

        private System.Windows.Forms.Timer _uiRefreshTimer;

        public SyncEngineForm()
        {
            InitializeComponent(); // This creates the UI
            LoadSettings();
            UpdateUIFromSettings();
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            // Initialize logging based on settings
            // Always enable during setup, disable after first successful sync (unless manually enabled)
            if (_settings.HasCompletedFirstSync && !_settings.EnableLogging)
            {
                Logger.SetEnabled(false);
            }
            else
            {
                Logger.SetEnabled(true);
                _settings.EnableLogging = true; // Ensure it's enabled during setup
            }

            // If settings are valid, start the service and hide the form.
            // If settings are invalid (first run), the form will be shown.
            if (IsSettingsValid())
            {
                this.Visible = false;
                this.ShowInTaskbar = false;
                StartSyncService();
            }
            else
            {
                this.Visible = true;
                this.ShowInTaskbar = true;
            }

            // ### Start UI refresh timer for updating "XX mins ago" ###
            _uiRefreshTimer = new System.Windows.Forms.Timer();
            _uiRefreshTimer.Interval = 60000; // 1 minute
            _uiRefreshTimer.Tick += (s, ev) => UpdateLastSyncLabel();
            _uiRefreshTimer.Start();
        }

        private bool IsSettingsValid()
        {
            return _settings != null &&
                  !string.IsNullOrEmpty(_settings.ExcelFilePath) &&
             !string.IsNullOrEmpty(_settings.GoogleRefreshToken) &&
                 !string.IsNullOrEmpty(_settings.GoogleCalendarId);
        }

        private void LoadSettings()
        {
            try
            {
                if (File.Exists(SettingsFileName))
                {
                    var json = File.ReadAllText(SettingsFileName);
                    _settings = JsonConvert.DeserializeObject<SyncSettings>(json);
                }
            }
            catch (Exception ex)
            {
                Logger.Error("Error loading settings", ex);
            }

            if (_settings == null)
            {
                _settings = new SyncSettings();
            }
        }

        private void SaveSettings()
        {
            try
            {
                var json = JsonConvert.SerializeObject(_settings, Formatting.Indented);
                File.WriteAllText(SettingsFileName, json);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to save settings: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ### Populates all UI fields based on loaded settings ###
        private void UpdateUIFromSettings()
        {
            // --- Excel Panel ---
      lblExcelFile.Text = string.IsNullOrEmpty(_settings.ExcelFilePath)
              ? "No file selected."
       : Path.GetFileName(_settings.ExcelFilePath);
   ValidateExcelFile(); // This updates the "tick"

    // --- Google Panel ---
            if (!string.IsNullOrEmpty(_settings.GoogleRefreshToken))
          {
           // Logged-in state
      lblGoogleAccount.Text = "Status: Logged In";
         
     // Show email if available, otherwise show a helpful message
     if (!string.IsNullOrEmpty(_settings.GoogleAccountEmail))
   {
        lblGoogleCalendar.Text = $"Account: {_settings.GoogleAccountEmail}";
                }
           else
       {
   lblGoogleCalendar.Text = "Account: (email not saved - syncing still works)";
                }
   
    // Show calendar name if available
if (!string.IsNullOrEmpty(_settings.GoogleCalendarName))
    {
           lblCalendarName.Text = $"Calendar: {_settings.GoogleCalendarName}";
    }
           else if (!string.IsNullOrEmpty(_settings.GoogleCalendarId))
           {
                // Show ID if name is missing but ID exists
       lblCalendarName.Text = $"Calendar: {_settings.GoogleCalendarId}";
        }
      else
      {
    lblCalendarName.Text = "Calendar: (connected - syncing works)";
        }
  
            lblGoogleAccount.Visible = true;
 lblGoogleCalendar.Visible = true;
    lblCalendarName.Visible = true;
    btnLoginGoogle.Visible = false;
                btnLogoutGoogle.Visible = true;
       cmbGoogleCalendar.Visible = false;
    }
   else
            {
                // Logged-out state
         lblGoogleAccount.Text = "Status: Not logged in";
     lblGoogleCalendar.Text = "Account: None";
         lblCalendarName.Text = "Calendar: None";
     lblGoogleAccount.Visible = true;
  lblGoogleCalendar.Visible = true;
         lblCalendarName.Visible = true;
         btnLoginGoogle.Visible = true;
    btnLogoutGoogle.Visible = false;
          cmbGoogleCalendar.Visible = false; // Hide until login
      }

        // --- Sync Panel ---
  numSyncInterval.Value = Math.Max(1, _settings.SyncIntervalMinutes); // Ensure at least 1
            chkEnableLogging.Checked = _settings.EnableLogging; // Reflect current logging state
     UpdateLastSyncLabel(); // Update last sync display
        }

        // ### Update last sync label with relative time ###
        private void UpdateLastSyncLabel()
        {
            if (!_settings.LastSyncTime.HasValue)
            {
                lblLastSync.Text = "Last sync: Never";
                lblLastSync.ForeColor = Color.Gray;
                return;
            }

            var elapsed = DateTime.Now - _settings.LastSyncTime.Value;
            string timeAgo;

            if (elapsed.TotalMinutes < 1)
                timeAgo = "just now";
            else if (elapsed.TotalMinutes < 60)
                timeAgo = $"{(int)elapsed.TotalMinutes} min ago";
            else if (elapsed.TotalHours < 24)
                timeAgo = $"{(int)elapsed.TotalHours} hr ago";
            else
                timeAgo = $"{(int)elapsed.TotalDays} day(s) ago";

            lblLastSync.Text = $"Last sync: {timeAgo} ({_settings.LastSyncTime.Value:g})";
            lblLastSync.ForeColor = elapsed.TotalMinutes < _settings.SyncIntervalMinutes * 2
         ? Color.Green
            : Color.Orange; // Warn if sync hasn't run in 2x the interval
        }

        private void ValidateExcelFile()
        {
            if (string.IsNullOrEmpty(_settings.ExcelFilePath))
   {
       lblExcelStatus.Text = "ERROR";
    lblExcelStatus.ForeColor = Color.Red;
    return;
    }

            if (!File.Exists(_settings.ExcelFilePath))
            {
   lblExcelStatus.Text = "ERROR";
      lblExcelStatus.ForeColor = Color.Red;
            return;
            }

            try
            {
                // Test read. We just need to check headers, not read all data.
                using (var workbook = new XLWorkbook(_settings.ExcelFilePath))
        {
  var worksheet = workbook.Worksheets.FirstOrDefault();
  if (worksheet == null) throw new Exception("No worksheet found.");
          var table = worksheet.Table(ExcelTableName);
if (table == null) throw new Exception($"Table '{ExcelTableName}' not found.");

 var headers = table.HeadersRow().Cells().Select(c => c.GetValue<string>().Trim()).ToList();
         string[] requiredHeaders = { "IsProcessed", "Outlook Event ID", "ActionType" };
          foreach (var reqHeader in requiredHeaders)
        {
 if (!headers.Any(h => h.Equals(reqHeader, StringComparison.OrdinalIgnoreCase)))
              {
 throw new Exception($"Missing header: {reqHeader}");
  }
        }
                }

       // If all checks pass:
   lblExcelStatus.Text = "OK";
      lblExcelStatus.ForeColor = Color.Green;
   }
            catch (Exception ex)
            {
  // File is locked or invalid
        lblExcelStatus.Text = "ERROR";
             lblExcelStatus.ForeColor = Color.Red;
          Logger.Warning($"Excel validation failed: {ex.Message}");
      }
        }

        private void StartSyncService()
        {
            InitializeTrayIcon();
            InitializeIdMap();
            InitializeSyncTimer(); 

            // Run a sync on startup - consider delaying if loading at windows start
            TriggerSync();
        }

        private void InitializeSyncTimer()
        {
            if (_syncTimer != null)
            {
                _syncTimer.Dispose();
            }

            int intervalMs = _settings.SyncIntervalMinutes * 60 * 1000;
            _syncTimer = new System.Threading.Timer(OnTimerElapsed, null, intervalMs, intervalMs);
            Logger.Info($"Sync timer started. Interval: {_settings.SyncIntervalMinutes} minutes.");
        }

        private void OnTimerElapsed(object state)
        {
            Logger.Info("Timer elapsed. Triggering sync...");
            TriggerSync();
        }

        // ### TRAY APP & SYNC LOGIC (Mostly Unchanged) ###

        private void InitializeTrayIcon()
        {
            if (trayIcon != null) return; // Already initialized

            trayMenu = new ContextMenuStrip();
            trayMenu.Items.Add("Sync Now", null, OnSyncNow);
            trayMenu.Items.Add("Show Settings", null, OnShowSettings);
            trayMenu.Items.Add("Open Logs Folder", null, (s, ev) => Logger.OpenLogDirectory());
            trayMenu.Items.Add(new ToolStripSeparator());
            trayMenu.Items.Add("Exit", null, OnExit);

            trayIcon = new NotifyIcon
            {
                Text = "Phil's Super Syncer (Running)",
                Icon = LoadIconFromResource("SyncOutlookToGoogle.icon.ico"),
                ContextMenuStrip = trayMenu,
                Visible = true
            };

            trayIcon.DoubleClick += OnShowSettings;
        }

        // ### Load icon from embedded resource ###
        private Icon LoadIconFromResource(string resourceName)
        {
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                using (var stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream != null)
                    {
                        return new Icon(stream);
                    }
                    else
                    {
                        Logger.Warning($"Icon resource '{resourceName}' not found. Available resources:");
                        foreach (var name in assembly.GetManifestResourceNames())
                        {
                            Logger.Info($"  - {name}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Failed to load icon from resource: {ex.Message}");
            }

            // Fallback to system icon
            Logger.Warning("Using fallback system icon");
            return SystemIcons.Application;
        }

        private void OnShowSettings(object sender, EventArgs e)
        {
            this.Visible = true;
            this.ShowInTaskbar = true;
            this.Activate();
            UpdateUIFromSettings(); // Refresh UI when shown
        }

        // ### REPLACED DATABASE WITH ID MAP ###
        private void InitializeIdMap()
        {
            Logger.Info("Initializing ID Map...");
            try
            {
                lock (_idMapLock)
                {
                    if (File.Exists(IdMapFileName))
                    {
                        var json = File.ReadAllText(IdMapFileName);
                        _idMap = JsonConvert.DeserializeObject<Dictionary<string, string>>(json);
                        Logger.Info($"ID Map loaded with {_idMap.Count} entries.");
                    }
                    else
                    {
                        _idMap = new Dictionary<string, string>();
                        Logger.Info("No ID Map found. Creating new one.");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"Failed to initialize ID Map: {ex.Message}");
                _idMap = new Dictionary<string, string>(); // Start fresh if file is corrupt
            }
        }

        // ### NEW METHOD TO SAVE THE ID MAP ###
        private void SaveIdMap()
        {
            Logger.Info("Saving ID Map...");
            try
            {
                lock (_idMapLock)
                {
                    var json = JsonConvert.SerializeObject(_idMap, Formatting.Indented);
                    File.WriteAllText(IdMapFileName, json);
                }
                Logger.Info("ID Map saved.");
            }
            catch (Exception ex)
            {
                Logger.Error($"Failed to save ID Map: {ex.Message}");
            }
        }

        private void InitializeFileWatcher()
        {
            try
            {
                string folder = Path.GetDirectoryName(_settings.ExcelFilePath);
                string fileName = Path.GetFileName(_settings.ExcelFilePath);

                fileWatcher = new FileSystemWatcher(folder)
                {
                    Filter = fileName,
                    NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName,
                    EnableRaisingEvents = true
                };

                fileWatcher.Changed += OnExcelFileChanged;
                fileWatcher.Created += OnExcelFileChanged;
                fileWatcher.Renamed += OnExcelFileRenamed;

                Logger.Info($"File watcher started for: {_settings.ExcelFilePath}");
            }
            catch (Exception ex)
            {
                Logger.Error($"Error starting file watcher: {ex.Message}");
                if (trayIcon != null)
                {
                    trayIcon.ShowBalloonTip(5000, "Sync Engine Error", $"Error starting file watcher: {ex.Message}", ToolTipIcon.Error);
                }
            }
        }

        private void OnExcelFileRenamed(object sender, RenamedEventArgs e)
        {
            if (e.OldName == Path.GetFileName(_settings.ExcelFilePath))
            {
                if (fileWatcher != null) fileWatcher.EnableRaisingEvents = false;
                Logger.Info("File renamed away. Watcher paused.");
            }
            if (e.Name == Path.GetFileName(_settings.ExcelFilePath))
            {
                if (fileWatcher != null) fileWatcher.EnableRaisingEvents = true;
                Logger.Info("File renamed to target. Watcher (re)started.");
                TriggerSync();
            }
        }

        private void OnExcelFileChanged(object sender, FileSystemEventArgs e)
        {
            Logger.Info($"File change detected: {e.ChangeType}");
            TriggerSync();
        }

        private void OnSyncNow(object sender, EventArgs e)
        {
            Logger.Info("Manual sync triggered.");
            TriggerSync();
        }

        private void TriggerSync()
        {
            if (_isSyncing)
            {
                Logger.Info("Sync already in progress. Skipping trigger.");
                return;
            }

            lock (_syncLock)
            {
                if (_isSyncing) return;
                _isSyncing = true;
            }

            Task.Run(async () =>
           {
               try
               {
                   Thread.Sleep(3000);
                   await ProcessExcelFile();
               }
               catch (Exception ex)
               {
                   Logger.Info($"Sync failed: {ex.Message}");
               }
               finally
               {
                   _isSyncing = false;
               }
           });
        }

        // ### TODO: FLaky, no robustness ###
        private async Task ProcessExcelFile()
        {
            Logger.Info("Starting Excel file processing...");

            // 1. Get Google Calendar Service
            CalendarService service;
            try
            {
                service = await GoogleAuth.GetCalendarServiceAsync(_settings.GoogleRefreshToken);
            }
            catch (Exception ex)
            {
                Logger.Error($"Failed to get Google service: {ex.Message}");
                if (trayIcon != null)
                {
                    trayIcon.ShowBalloonTip(3000, "Sync Error", "Failed to authenticate with Google. Please reset and log in.", ToolTipIcon.Error);
                }
                return;
            }

            // 2. Open Excel File into memory to avoid file locking
            MemoryStream workbookStream;
            try
            {
                byte[] fileBytes = File.ReadAllBytes(_settings.ExcelFilePath);
                workbookStream = new MemoryStream(fileBytes);
                Logger.Info("Excel file read into memory.");
            }
            catch (IOException ioEx)
            {
                Logger.Info($"Excel file is likely open or in use (read failed): {ioEx.Message}");
                if (trayIcon != null)
                {
                    trayIcon.ShowBalloonTip(3000, "Sync Error", "Excel file is open or in use. Close it and try again.", ToolTipIcon.Warning);
                }
                return;
            }
            catch (Exception ex)
            {
                Logger.Error($"Failed to read Excel file: {ex.Message}");
                if (trayIcon != null)
                {
                    trayIcon.ShowBalloonTip(3000, "Sync Error", $"Failed to read Excel: {ex.Message}", ToolTipIcon.Error);
                }
                return;
            }

            // 3. Process the workbook from memory
            try
            {
                using (var workbook = new XLWorkbook(workbookStream))
                {
                    var worksheet = workbook.Worksheets.FirstOrDefault();
                    if (worksheet == null) throw new Exception("No worksheet found in Excel file.");

                    var table = worksheet.Table(ExcelTableName);
                    if (table == null) throw new Exception($"Table '{ExcelTableName}' not found. Did you format it as a table?");

                    // ### Map headers to their column index (e.g., "IsProcessed" -> 7) ###
                    var headers = table.HeadersRow().Cells().Select(c => c.GetValue<string>().Trim()).ToList();
                    var columnIndexes = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                    int colIdx = 1;
                    foreach (var header in headers)
                    {
                        if (!columnIndexes.ContainsKey(header))
                        {
                            columnIndexes[header] = colIdx;
                        }
                        colIdx++;
                    }
                    Logger.Info("Found table headers: " + string.Join(", ", headers));

                    // Check for our required headers
                    string[] requiredHeaders = { "IsProcessed", "Outlook Event ID", "ActionType", "Timestamp", "isAllDay", "Subject", "Body", "Location", "StartTime", "EndTime" };
                    foreach (var reqHeader in requiredHeaders)
                    {
                        if (!columnIndexes.ContainsKey(reqHeader))
                        {
                            Logger.Error($"CRITICAL ERROR: Header '{reqHeader}' not found in Excel file.");
                            Logger.Info("Please ensure your Excel headers match the code exactly.");
                            throw new Exception($"Required header '{reqHeader}' not found in Excel file. Found headers: [{string.Join(", ", headers)}]");
                        }
                    }
                    Logger.Info("All required headers found. Proceeding...");

                    // ### MODIFIED: Use column index instead of name & PARSE STRING ###
                    var unprocessedRows = table.DataRange.Rows()
                .Where(r => r.Cell(columnIndexes["IsProcessed"]).GetValue<string>().Equals("FALSE", StringComparison.OrdinalIgnoreCase))
                          .OrderBy(r => DateTimeOffset.Parse(r.Cell(columnIndexes["Timestamp"]).GetValue<string>()))
                 .ToList();

                    if (unprocessedRows.Count == 0)
                    {
                        Logger.Info("No new events to process.");
                        workbookStream.Dispose();

                        // Update last sync time even when no events to process
                        _settings.LastSyncTime = DateTime.Now;
                        SaveSettings();

                        // Update UI on main thread
                        if (this.InvokeRequired)
                        {
                            this.BeginInvoke(new Action(() => UpdateLastSyncLabel()));
                        }
                        else
                        {
                            UpdateLastSyncLabel();
                        }

                        return;
                    }

                    Logger.Info($"Found {unprocessedRows.Count} new event(s) to process.");
                    if (trayIcon != null)
                    {
                        trayIcon.ShowBalloonTip(1000, "Syncing...", $"Syncing {unprocessedRows.Count} new event(s)...", ToolTipIcon.Info);
                    }

                    bool changesMadeToMap = false;

                    // 4. Loop Through Rows
                    foreach (var row in unprocessedRows)
                    {
                        try
                        {
                            string outlookId = row.Cell(columnIndexes["Outlook Event ID"]).GetValue<string>();
                            string actionType = row.Cell(columnIndexes["ActionType"]).GetValue<string>();

                            bool success = false;
                            switch (actionType.ToLower())
                            {
                                case "added":
                                    success = await HandleAddedEvent(service, _idMap, row, outlookId, columnIndexes);
                                    if (success) changesMadeToMap = true;
                                    break;
                                case "updated":
                                    success = await HandleUpdatedEvent(service, _idMap, row, outlookId, columnIndexes);
                                    if (success) changesMadeToMap = true;
                                    break;
                                case "deleted":
                                    success = await HandleDeletedEvent(service, _idMap, row, outlookId, columnIndexes);
                                    if (success) changesMadeToMap = true;
                                    break;
                            }

                            // 5. Mark as Processed
                            if (success)
                            {
                                row.Cell(columnIndexes["IsProcessed"]).Value = "TRUE";
                            }
                        }
                        catch (Exception ex)
                        {
                            Logger.Error($"Failed to process row {row.RowNumber()}: {ex.Message}");
                        }
                    }

                    // 6. Save Excel (only if changes were made)
                    if (unprocessedRows.Count > 0)
                    {
                        using (var ms = new MemoryStream())
                        {
                            workbook.SaveAs(ms);
                            File.WriteAllBytes(_settings.ExcelFilePath, ms.ToArray());
                        }
                        Logger.Info("Sync complete. Excel file updated.");
                    }

                    // 7. Save the ID map if it changed
                    if (changesMadeToMap)
                    {
                        SaveIdMap();
                    }

                    // 8. Update last sync time
                    _settings.LastSyncTime = DateTime.Now;
   
        // Mark first sync as complete and auto-disable logging (unless manually enabled)
        if (!_settings.HasCompletedFirstSync)
  {
         _settings.HasCompletedFirstSync = true;
     _settings.EnableLogging = false; // Auto-disable after first successful sync
       Logger.Info("=== First successful sync completed. Logging will now be disabled by default. ===");
Logger.Info("You can re-enable logging in Settings if needed.");
Logger.SetEnabled(false);
            }
    
    SaveSettings();

            // 9. Update UI on main thread
            if (this.InvokeRequired)
            {
                this.BeginInvoke(new Action(() => UpdateLastSyncLabel()));
            }
            else
            {
                UpdateLastSyncLabel();
            }
                }
                workbookStream.Dispose();
            }
            catch (Exception ex)
            {
                Logger.Error($"Failed to process Excel file: {ex.Message}");
                if (trayIcon != null)
                {
                    trayIcon.ShowBalloonTip(3000, "Sync Error", $"Failed to process Excel: {ex.Message}", ToolTipIcon.Error);
                }
            }
        }

        // ### SYNC HELPER METHODS (UPDATED TO USE DICTIONARY & INDEXES) ###

        private async Task<bool> HandleAddedEvent(CalendarService service, Dictionary<string, string> idMap, IXLTableRow row, string outlookId, Dictionary<string, int> columnIndexes)
        {
            Logger.Info($"Handling 'added' for Outlook ID: {outlookId}");
            var newEvent = CreateEventFromRow(row, columnIndexes);

            var insertRequest = service.Events.Insert(newEvent, _settings.GoogleCalendarId);
            var createdEvent = await insertRequest.ExecuteAsync();

            if (createdEvent != null)
            {
                lock (_idMapLock)
                {
                    idMap[outlookId] = createdEvent.Id;
                }
                Logger.Info($"Created Google event {createdEvent.Id}");
                return true;
            }
            return false;
        }

        private async Task<bool> HandleUpdatedEvent(CalendarService service, Dictionary<string, string> idMap, IXLTableRow row, string outlookId, Dictionary<string, int> columnIndexes)
        {
            Logger.Info($"Handling 'updated' for Outlook ID: {outlookId}");

            string googleId;
            lock (_idMapLock)
            {
                idMap.TryGetValue(outlookId, out googleId);
            }

            if (string.IsNullOrEmpty(googleId))
            {
                Logger.Info("...Update found no matching Google ID. Treating as 'add'.");
                return await HandleAddedEvent(service, idMap, row, outlookId, columnIndexes);
            }

            var updatedEvent = CreateEventFromRow(row, columnIndexes);
            var updateRequest = service.Events.Update(updatedEvent, _settings.GoogleCalendarId, googleId);
            await updateRequest.ExecuteAsync();

            Logger.Info($"Updated Google event {googleId}");
            return true;
        }

        private async Task<bool> HandleDeletedEvent(CalendarService service, Dictionary<string, string> idMap, IXLTableRow row, string outlookId, Dictionary<string, int> columnIndexes)
        {
            Logger.Info($"Handling 'deleted' for Outlook ID: {outlookId}");

            string googleId;
            lock (_idMapLock)
            {
                idMap.TryGetValue(outlookId, out googleId);
            }

            if (string.IsNullOrEmpty(googleId))
            {
                Logger.Info("...Delete found no matching Google ID. Skipping.");
                return true;
            }

            try
            {
                var deleteRequest = service.Events.Delete(_settings.GoogleCalendarId, googleId);
                await deleteRequest.ExecuteAsync();
                Logger.Info($"Deleted Google event {googleId}");
            }
            catch (Google.GoogleApiException gEx)
            {
                if (gEx.HttpStatusCode == System.Net.HttpStatusCode.NotFound)
                {
                    Logger.Info("...Google event was already deleted. Ignoring.");
                }
                else
                {
                    throw;
                }
            }

            lock (_idMapLock)
            {
                idMap.Remove(outlookId);
            }
            return true;
        }

        // ### HELPER METHODS ###

        private Event CreateEventFromRow(IXLTableRow row, Dictionary<string, int> columnIndexes)
        {
            var newEvent = new Event();
            newEvent.Summary = row.Cell(columnIndexes["Subject"]).GetValue<string>();
            newEvent.Description = row.Cell(columnIndexes["Body"]).GetValue<string>();
            newEvent.Location = row.Cell(columnIndexes["Location"]).GetValue<string>();

            string isAllDayString = row.Cell(columnIndexes["isAllDay"]).GetValue<string>();
            bool isAllDay = false;
            bool.TryParse(isAllDayString, out isAllDay);

            string startTimeString = row.Cell(columnIndexes["StartTime"]).GetValue<string>();
            string endTimeString = row.Cell(columnIndexes["EndTime"]).GetValue<string>();

            DateTimeOffset startTime = DateTimeOffset.Parse(startTimeString);
            DateTimeOffset endTime = DateTimeOffset.Parse(endTimeString);

            if (isAllDay)
            {
                newEvent.Start = new EventDateTime { Date = startTime.ToString("yyyy-MM-dd") };
                newEvent.End = new EventDateTime { Date = endTime.ToString("yyyy-MM-dd") };
            }
            else
            {
                newEvent.Start = new EventDateTime { DateTimeDateTimeOffset = startTime };
                newEvent.End = new EventDateTime { DateTimeDateTimeOffset = endTime };
            }

            return newEvent;
        }

        // ### FORM & APPLICATION LIFECYCLE  ###

        private void OnExit(object sender, EventArgs e)
        {
            if (trayIcon != null)
            {
                trayIcon.Visible = false;
            }
            Application.Exit();
        }

        private System.ComponentModel.IContainer components = null;
        private FileSystemWatcher fileWatcher;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            if (disposing)
            {
                if (trayIcon != null) trayIcon.Dispose();
                if (trayMenu != null) trayMenu.Dispose();
                if (fileWatcher != null) fileWatcher.Dispose();
                if (_syncTimer != null) _syncTimer.Dispose();
                if (_uiRefreshTimer != null) _uiRefreshTimer.Dispose();
            }
            base.Dispose(disposing);
        }

        private void OnBrowseExcel(object sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                dialog.Title = "Select your sync log file";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    _settings.ExcelFilePath = dialog.FileName;
                    SaveSettings();
                    UpdateUIFromSettings();
                }
            }
        }

        private async void OnGoogleLogin(object sender, EventArgs e)
        {
            try
            {
                btnLoginGoogle.Enabled = false;
                lblGoogleAccount.Text = "Status: Waiting for login...";

                var (credential, calendars, email) = await GoogleAuth.AuthorizeAsync();

                // Save the email (important!)
                  _settings.GoogleAccountEmail = email;
                 Logger.Info($"Google login successful. Email: {email ?? "NULL"}");

              if (credential.Token.RefreshToken != null)
    {
       _settings.GoogleRefreshToken = credential.Token.RefreshToken;
          Logger.Info("Refresh token obtained from credential.");
     }
    else
          {
          Logger.Warning("Refresh token was null in credential.Token. Attempting to read from token file...");
   string tokenFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "token.json");
           string tokenFile = Path.Combine(tokenFolder, "Google.Apis.Auth.OAuth2.Responses.TokenResponse-user");
  if (File.Exists(tokenFile))
{
                 var tokenJson = File.ReadAllText(tokenFile);
                var token = JsonConvert.DeserializeObject<Google.Apis.Auth.OAuth2.Responses.TokenResponse>(tokenJson);
_settings.GoogleRefreshToken = token.RefreshToken;
     Logger.Info("Refresh token obtained from token file.");
}
        else
     {
           Logger.Error("Token file not found. Refresh token could not be obtained.");
}
           }

            cmbGoogleCalendar.Items.Clear();
           foreach (var calendar in calendars.Items)
         {
              cmbGoogleCalendar.Items.Add(new CalendarItem { Id = calendar.Id, Summary = calendar.Summary });
    }
             cmbGoogleCalendar.DisplayMember = "Summary";
      cmbGoogleCalendar.ValueMember = "Id";

     // Pre-select primary calendar
     var primary = calendars.Items.FirstOrDefault(c => c.Primary == true);
    if (primary != null)
 {
         cmbGoogleCalendar.SelectedItem = cmbGoogleCalendar.Items
           .Cast<CalendarItem>()
        .FirstOrDefault(ci => ci.Id == primary.Id);
                }

  // Show the dropdown and ask user to select
           cmbGoogleCalendar.Visible = true;
   lblGoogleAccount.Text = "Status: Logged In";
    lblGoogleCalendar.Text = $"Account: {email ?? "Unknown"}";
          lblCalendarName.Text = "Please select a calendar below:";

      // Hide login button, show logout button
      btnLoginGoogle.Visible = false;
         btnLogoutGoogle.Visible = true;

           // IMPORTANT: Save settings immediately after login
         SaveSettings();
     Logger.Info($"Settings saved after login. Email saved: {_settings.GoogleAccountEmail}");
     }
            catch (Exception ex)
 {
      Logger.Error($"Google login failed: {ex.Message}", ex);
      MessageBox.Show($"Google login failed: {ex.Message}", "Login Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
     UpdateUIFromSettings();
       }
            finally
          {
      btnLoginGoogle.Enabled = true;
    }
}

        private void OnGoogleCalendarSelected(object sender, EventArgs e)
   {
   var selectedCalendar = cmbGoogleCalendar.SelectedItem as CalendarItem;
 if (selectedCalendar != null)
    {
        _settings.GoogleCalendarId = selectedCalendar.Id;
     _settings.GoogleCalendarName = selectedCalendar.Summary;
   
     Logger.Info($"Calendar selected: {selectedCalendar.Summary} (ID: {selectedCalendar.Id})");

  SaveSettings();
    Logger.Info($"Settings saved after calendar selection. Email: {_settings.GoogleAccountEmail}, Calendar: {_settings.GoogleCalendarName}");

       // Ensure the calendar dropdown is hidden after selection
    cmbGoogleCalendar.Visible = false;

      // Update UI to reflect the changes
      UpdateUIFromSettings();

    if (IsSettingsValid() && _syncTimer == null)
     {
    StartSyncService();
 }
   }
        }

        private void OnGoogleLogout(object sender, EventArgs e)
        {
            if (_syncTimer != null)
            {
                _syncTimer.Dispose();
                _syncTimer = null;
            }

            _settings.GoogleRefreshToken = null;
            _settings.GoogleAccountEmail = null;
            _settings.GoogleCalendarId = null;
            _settings.GoogleCalendarName = null;
            SaveSettings();

            try
            {
                string tokenFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "token.json");
                if (Directory.Exists(tokenFolder))
                {
                    Directory.Delete(tokenFolder, true);
                }
            }
            catch (Exception ex)
            {
                Logger.Info($"Could not delete token folder: {ex.Message}");
            }

            UpdateUIFromSettings();
        }

        private void OnSave(object sender, EventArgs e)
        {
            _settings.SyncIntervalMinutes = (int)numSyncInterval.Value;

   // Update logging preference
            bool loggingChanged = _settings.EnableLogging != chkEnableLogging.Checked;
            _settings.EnableLogging = chkEnableLogging.Checked;

            SaveSettings();

            if (loggingChanged)
          {
   Logger.SetEnabled(_settings.EnableLogging);
            }

            if (_syncTimer != null)
    {
          InitializeSyncTimer();
    }

          // Show confirmation and minimize to tray
    MessageBox.Show("Settings saved successfully!\n\nThe app will continue running in the system tray.", 
                "Settings Saved", 
    MessageBoxButtons.OK, 
      MessageBoxIcon.Information);

  this.Hide();
 
   // Show tray notification
   if (trayIcon != null)
   {
     trayIcon.ShowBalloonTip(2000, "Phil's Super Syncer", "Settings saved. App is running in the system tray.", ToolTipIcon.Info);
    }
        }

        private void OnResetApp(object sender, EventArgs e)
        {
            var result = MessageBox.Show(
          "This will log you out, delete ALL settings, and clear the event tracking history (id_map.json).\n\nThe application will restart.\n\nAre you sure you want to reset?",
     "Reset Confirmation",
              MessageBoxButtons.YesNo,
       MessageBoxIcon.Warning
     );

            if (result == DialogResult.Yes)
            {
                if (_syncTimer != null) _syncTimer.Dispose();
                if (trayIcon != null) trayIcon.Dispose();

                try
                {
                    if (File.Exists(SettingsFileName)) File.Delete(SettingsFileName);
                    if (File.Exists(IdMapFileName)) File.Delete(IdMapFileName);

                    string tokenFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "token.json");
                    if (Directory.Exists(tokenFolder))
                    {
                        Directory.Delete(tokenFolder, true);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Could not fully clean up old files: {ex.Message}", "Reset Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                Application.Restart();
            }
        }

        [STAThread]
        static void Main()
        {
            Logger.Info("=== Phil's Super Syncer Starting ===");
            Logger.Info($"Version: {System.Reflection.Assembly.GetExecutingAssembly().GetName().Version}");
            Logger.Info($"Log file: {Logger.GetCurrentLogFilePath()}");

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new SyncEngineForm());

            Logger.Info("=== Phil's Super Syncer Exiting ===");
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(434, 491); // Increased from 461 to 491 (30px taller)
            this.Name = "SyncEngineForm";
            this.Text = "Phil's Super Syncer - Settings";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            // Set form icon
            this.Icon = LoadIconFromResource("SyncOutlookToGoogle.icon.ico");

            // --- Excel Group ---
            this.grpExcel = new GroupBox { Text = "1. Outlook / Excel File", Location = new Point(12, 12), Size = new Size(410, 90) };
            this.lblExcelFile = new Label { Text = "No file selected.", Location = new Point(15, 25), AutoSize = true, MaximumSize = new Size(280, 40) };
            this.lblExcelStatus = new Label { Text = "ERROR", Location = new Point(360, 25), Font = new Font(this.Font, FontStyle.Bold), ForeColor = Color.Red, AutoSize = true };
            this.btnChangeExcel = new Button { Text = "Change...", Location = new Point(320, 50), Size = new Size(75, 25) };
            this.btnChangeExcel.Click += OnBrowseExcel;

         this.grpExcel.Controls.Add(this.lblExcelFile);
     this.grpExcel.Controls.Add(this.lblExcelStatus);
      this.grpExcel.Controls.Add(this.btnChangeExcel);
            this.Controls.Add(this.grpExcel);

            // --- Google Group ---
            this.grpGoogle = new GroupBox { Text = "2. Google Calendar", Location = new Point(12, 110), Size = new Size(410, 160) };
            this.lblGoogleAccount = new Label { Text = "Status: Not logged in", Location = new Point(15, 25), AutoSize = true };
            this.lblGoogleCalendar = new Label { Text = "Account: None", Location = new Point(15, 50), AutoSize = true };
            this.lblCalendarName = new Label { Text = "Calendar: None", Location = new Point(15, 75), AutoSize = true };
            this.btnLoginGoogle = new Button { Text = "Login to Google", Location = new Point(15, 100), Size = new Size(150, 25) };
            this.btnLoginGoogle.Click += OnGoogleLogin;
            this.btnLogoutGoogle = new Button { Text = "Logout", Location = new Point(320, 25), Size = new Size(75, 25) };
            this.btnLogoutGoogle.Click += OnGoogleLogout;
            this.cmbGoogleCalendar = new ComboBox { Location = new Point(15, 130), Size = new Size(380, 23), DropDownStyle = ComboBoxStyle.DropDownList, Visible = false };
            this.cmbGoogleCalendar.SelectionChangeCommitted += OnGoogleCalendarSelected;

            this.grpGoogle.Controls.Add(this.lblGoogleAccount);
            this.grpGoogle.Controls.Add(this.lblGoogleCalendar);
            this.grpGoogle.Controls.Add(this.lblCalendarName);
            this.grpGoogle.Controls.Add(this.btnLoginGoogle);
            this.grpGoogle.Controls.Add(this.btnLogoutGoogle);
            this.grpGoogle.Controls.Add(this.cmbGoogleCalendar);
            this.Controls.Add(this.grpGoogle);

            // --- Sync Group ---
            this.grpSync = new GroupBox { Text = "3. Sync Settings", Location = new Point(12, 280), Size = new Size(410, 120) }; // Increased height
            this.lblSyncInterval = new Label { Text = "Check for changes every (minutes):", Location = new Point(15, 25), AutoSize = true };
 this.numSyncInterval = new NumericUpDown { Location = new Point(220, 23), Size = new Size(60, 23), Minimum = 1, Maximum = 120 };
            this.lblLastSync = new Label { Text = "Last sync: Never", Location = new Point(15, 55), AutoSize = true, ForeColor = Color.Gray };
  this.btnSyncNow = new Button { Text = "Sync Now", Location = new Point(320, 50), Size = new Size(75, 25) };
            this.btnSyncNow.Click += OnSyncNow;
       this.chkEnableLogging = new CheckBox { Text = "Enable detailed logging", Location = new Point(15, 85), AutoSize = true };
            this.chkEnableLogging.CheckedChanged += (s, ev) =>
     {
        // Real-time update of logging state (without saving)
                if (!this.chkEnableLogging.Checked && !_settings.HasCompletedFirstSync)
          {
 // Don't allow disabling during initial setup
       this.chkEnableLogging.Checked = true;
     MessageBox.Show("Logging cannot be disabled until the first successful sync is complete.", 
             "Logging Required", MessageBoxButtons.OK, MessageBoxIcon.Information);
  }
     };

     this.grpSync.Controls.Add(this.lblSyncInterval);
     this.grpSync.Controls.Add(this.numSyncInterval);
      this.grpSync.Controls.Add(this.lblLastSync);
    this.grpSync.Controls.Add(this.btnSyncNow);
        this.grpSync.Controls.Add(this.chkEnableLogging);
       this.Controls.Add(this.grpSync);

            // --- Main Buttons ---
            this.btnSave = new Button { Text = "Save", Location = new Point(12, 410), Size = new Size(133, 30) }; // Moved down from 377 to 410
  this.btnSave.Click += OnSave;
  this.Controls.Add(this.btnSave);

         this.btnResetApp = new Button { Text = "Reset All", Location = new Point(151, 410), Size = new Size(133, 30) }; // Moved down from 377 to 410
   this.btnResetApp.Click += OnResetApp;
 this.Controls.Add(this.btnResetApp);

    Button btnExitApp = new Button { Text = "Exit App", Location = new Point(290, 410), Size = new Size(132, 30) }; // Moved down from 377 to 410
     btnExitApp.Click += OnExit;
  this.Controls.Add(btnExitApp);
            // Hide form on close
            this.FormClosing += (s, e) =>
       {
           if (e.CloseReason == CloseReason.UserClosing)
           {
               e.Cancel = true;
               this.Hide();
           }
       };
        }
    }
}

