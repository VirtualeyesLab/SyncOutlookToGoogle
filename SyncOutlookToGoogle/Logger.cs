using System;
using System.Diagnostics;
using System.IO;
using System.Text;

namespace CalendarSyncEngine
{
	/// <summary>
	/// Simple file-based logger with timestamp support
	/// Logs are written to a daily log file in the application directory
 /// </summary>
	public static class Logger
	{
		private static readonly object _logLock = new object();
   private static string _logDirectory;
	 private static string _currentLogFile;
private static readonly int MaxLogSizeBytes = 5 * 1024 * 1024; // 5 MB
  private static readonly int MaxLogFiles = 10; // Keep last 10 log files
  private static bool _isEnabled = true; // Default: enabled

		static Logger()
	 {
	  // Log directory is in the same folder as the executable
	_logDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Logs");

 try
	 {
	  if (!Directory.Exists(_logDirectory))
		{
   Directory.CreateDirectory(_logDirectory);
}
	  }
 catch (Exception ex)
			{
   System.Diagnostics.Debug.WriteLine($"Failed to create log directory: {ex.Message}");
  }

   UpdateLogFileName();
	  CleanupOldLogs();
 }

	/// <summary>
   /// Enable or disable logging
	/// </summary>
	public static void SetEnabled(bool enabled)
 {
			lock (_logLock)
			{
	 _isEnabled = enabled;
		if (enabled)
  {
		Info("=== Logging Enabled ===");
	 }
		   else
	{
	   Info("=== Logging Disabled ===");
				}
	}
		}

   /// <summary>
		/// Check if logging is currently enabled
	 /// </summary>
   public static bool IsEnabled()
		{
			return _isEnabled;
		}

		/// <summary>
		/// Updates the log file name based on current date
	 /// </summary>
		private static void UpdateLogFileName()
		{
			string date = DateTime.Now.ToString("yyyy-MM-dd");
		_currentLogFile = Path.Combine(_logDirectory, $"PhilsSuperSyncer_{date}.log");
		}

		/// <summary>
   /// Removes old log files, keeping only the most recent ones
   /// </summary>
		private static void CleanupOldLogs()
{
			try
			{
	  var logFiles = Directory.GetFiles(_logDirectory, "PhilsSuperSyncer_*.log");
		 if (logFiles.Length > MaxLogFiles)
		   {
		Array.Sort(logFiles);
	   for (int i = 0; i < logFiles.Length - MaxLogFiles; i++)
		{
		try
	   {
	File.Delete(logFiles[i]);
	   }
	 catch
		  {
  // Ignore deletion errors
	   }
   }
		 }
  }
			catch
	  {
				// Ignore cleanup errors
			}
		}

/// <summary>
		/// Checks if current log file exceeds size limit and rotates if needed
		/// </summary>
 private static void RotateLogIfNeeded()
		{
	  try
	  {
 if (File.Exists(_currentLogFile))
		 {
		var fileInfo = new FileInfo(_currentLogFile);
		   if (fileInfo.Length > MaxLogSizeBytes)
	   {
	  string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HHmmss");
		   string rotatedFile = Path.Combine(_logDirectory, $"PhilsSuperSyncer_{timestamp}.log");
			   File.Move(_currentLogFile, rotatedFile);
	  }
	   }
 }
	 catch
	 {
		   // If rotation fails, continue logging to the same file
			}
		}

		/// <summary>
   /// Writes a log entry with timestamp and level
		/// </summary>
  private static void Write(string level, string message)
		{
			if (!_isEnabled) return; // Skip if logging is disabled

	lock (_logLock)
   {
		  try
	{
UpdateLogFileName();
	 RotateLogIfNeeded();

		  string timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
		   string logEntry = $"[{timestamp}] [{level}] {message}{Environment.NewLine}";

		// Write to file
		File.AppendAllText(_currentLogFile, logEntry, Encoding.UTF8);

   // Also write to Debug output for development
		System.Diagnostics.Debug.Write(logEntry);
	}
				catch (Exception ex)
   {
		 // Fallback to Debug output only if file logging fails
	  System.Diagnostics.Debug.WriteLine($"[LOGGER ERROR] Failed to write to log file: {ex.Message}");
	System.Diagnostics.Debug.WriteLine($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] [{level}] {message}");
	 }
			}
		}

		/// <summary>
		/// Logs an informational message
		/// </summary>
		public static void Info(string message)
   {
	Write("INFO", message);
   }

	 /// <summary>
		/// Logs a warning message
		/// </summary>
		public static void Warning(string message)
   {
	   Write("WARN", message);
	  }

		/// <summary>
		/// Logs an error message
/// </summary>
		public static void Error(string message)
		{
 Write("ERROR", message);
		}

  /// <summary>
		/// Logs an error with exception details
		/// </summary>
		public static void Error(string message, Exception ex)
		{
	var sb = new StringBuilder();
	   sb.AppendLine(message);
		sb.AppendLine($"Exception Type: {ex.GetType().Name}");
			sb.AppendLine($"Exception Message: {ex.Message}");
	  sb.AppendLine($"Stack Trace: {ex.StackTrace}");

		  if (ex.InnerException != null)
   {
	   sb.AppendLine($"Inner Exception: {ex.InnerException.Message}");
	   }

	   Write("ERROR", sb.ToString());
  }

   /// <summary>
/// Logs a debug message (only in Debug builds)
		/// </summary>
		[Conditional("DEBUG")]
   public static void Debug(string message)
		{
	   Write("DEBUG", message);
		}

		/// <summary>
		/// Gets the path to the current log file
		/// </summary>
		public static string GetCurrentLogFilePath()
		{
		 return _currentLogFile;
		}

		/// <summary>
   /// Gets the logs directory path
		/// </summary>
		public static string GetLogDirectory()
		{
			return _logDirectory;
	 }

		/// <summary>
  /// Opens the log directory in Windows Explorer
		/// </summary>
	  public static void OpenLogDirectory()
		{
 try
		{
		if (Directory.Exists(_logDirectory))
		{
			  Process.Start("explorer.exe", _logDirectory);
		}
	   }
	 catch (Exception ex)
	{
		Error("Failed to open log directory", ex);
	 }
   }
	}
}
