"""
PyQt5 GUI for Outlook to Google Calendar Sync Agent.
Provides UI for managing settings, authentication, and viewing sync history.
"""

import sys
import os
import json
from datetime import datetime
from typing import Optional

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QLabel, QLineEdit, QPushButton, QSpinBox, QComboBox,
    QFileDialog, QTextEdit, QStatusBar, QMessageBox,
    QGroupBox, QFormLayout, QSystemTrayIcon, QMenu, QAction, QStyle
)
from PyQt5.QtCore import Qt, QTimer, pyqtSignal, QObject
from PyQt5.QtGui import QFont

from agent_config import ConfigManager
from sync import get_google_service, list_google_calendars, load_outlook_snapshot, perform_sync


class SyncSignals(QObject):
    """Signals for sync events."""
    sync_completed = pyqtSignal(dict)  # sync stats
    sync_error = pyqtSignal(str)


class SyncSettingsApp(QMainWindow):
    """Main application window for sync settings and management."""
    
    def __init__(self):
        super().__init__()
        self.config_manager = ConfigManager()
        self.signals = SyncSignals()
        self.is_quitting = False
        self.has_shown_tray_hint = False
        self.tray_icon = None
        self.logo_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "PhilsSyncLogo.png")
        self.init_ui()
        self.load_settings()
        self.setup_system_tray()

        # Timer to refresh status
        self.status_timer = QTimer()
        self.status_timer.timeout.connect(self.refresh_status)
        self.status_timer.start(5000)  # Update every 5 seconds
    
    def init_ui(self):
        """Initialize the user interface."""
        from PyQt5.QtGui import QIcon, QPixmap
        self.setWindowTitle("Outlook to Google Sync - Settings")
        self.setGeometry(100, 100, 900, 700)

        # Set window icon to logo
        if os.path.exists(self.logo_path):
            self.setWindowIcon(QIcon(self.logo_path))

        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # Add logo at the top
        if os.path.exists(self.logo_path):
            logo_label = QLabel()
            logo_pixmap = QPixmap(self.logo_path)
            logo_label.setPixmap(logo_pixmap.scaledToHeight(100, Qt.SmoothTransformation))
            logo_label.setAlignment(Qt.AlignCenter)
            main_layout.addWidget(logo_label)

        # Create tab widget
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)

        # Add tabs
        self.tabs.addTab(self.create_settings_tab(), "Settings")
        self.tabs.addTab(self.create_auth_tab(), "Google Authentication")
        self.tabs.addTab(self.create_history_tab(), "Sync History")

        # Status bar
        self.statusBar = QStatusBar()
        self.setStatusBar(self.statusBar)
        self.status_label = QLabel("Status: Ready")
        self.statusBar.addWidget(self.status_label, 1)
    
    def create_settings_tab(self) -> QWidget:
        """Create the settings tab."""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # Create form layout for settings
        form_layout = QFormLayout()
        
        # Outlook snapshot file path
        path_layout = QHBoxLayout()
        self.outlook_path_input = QLineEdit()
        self.outlook_path_input.setReadOnly(True)
        browse_btn = QPushButton("Browse...")
        browse_btn.clicked.connect(self.browse_outlook_file)
        path_layout.addWidget(self.outlook_path_input)
        path_layout.addWidget(browse_btn)
        form_layout.addRow("Outlook Snapshot File:", path_layout)
        
        # Timezone
        self.timezone_combo = QComboBox()
        timezones = [
            'Pacific/Auckland',
            'UTC',
            'US/Pacific',
            'US/Eastern',
            'Europe/London',
            'Europe/Berlin',
            'Asia/Tokyo',
            'Australia/Sydney'
        ]
        self.timezone_combo.addItems(timezones)
        form_layout.addRow("Timezone:", self.timezone_combo)
        
        # Sync frequency
        self.sync_frequency_spin = QSpinBox()
        self.sync_frequency_spin.setMinimum(1)
        self.sync_frequency_spin.setMaximum(1440)
        self.sync_frequency_spin.setSuffix(" minutes")
        form_layout.addRow("Sync Frequency:", self.sync_frequency_spin)
        
        # Logging level
        self.logging_level_combo = QComboBox()
        self.logging_level_combo.addItems(['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL'])
        form_layout.addRow("Logging Level:", self.logging_level_combo)

        # Selected target calendar (configured in auth tab)
        self.selected_calendar_label = QLabel("Primary Calendar (primary)")
        form_layout.addRow("Target Google Calendar:", self.selected_calendar_label)
        
        # Monitoring enabled
        monitoring_layout = QHBoxLayout()
        self.monitoring_check_label = QLabel("✓ File monitoring enabled")
        self.monitoring_check_label.setStyleSheet("color: green; font-weight: bold;")
        monitoring_layout.addWidget(self.monitoring_check_label)
        monitoring_layout.addStretch()
        form_layout.addRow("Monitoring:", monitoring_layout)
        
        layout.addLayout(form_layout)
        
        # Remove Save button (autosave now)

        close_note = QLabel("Tip: Minimize keeps the window in the taskbar. Close hides it to the system tray (monitoring continues). Use the tray icon to restore or quit.")
        close_note.setWordWrap(True)
        close_note.setStyleSheet("color: #555;")
        layout.addWidget(close_note)
        # Autosave connections for settings
        self.outlook_path_input.textChanged.connect(lambda: self.autosave_setting('OUTLOOK_JSON_PATH', self.outlook_path_input.text()))
        self.timezone_combo.currentTextChanged.connect(lambda tz: self.autosave_setting('TIMEZONE', tz))
        self.sync_frequency_spin.valueChanged.connect(lambda val: self.autosave_setting('SYNC_FREQUENCY_MINUTES', val))
        self.logging_level_combo.currentTextChanged.connect(lambda lvl: self.autosave_setting('LOGGING_LEVEL', lvl))
        
        layout.addStretch()
        return widget
    
    def create_auth_tab(self) -> QWidget:
        """Create the Google authentication tab."""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # Title
        title = QLabel("Google Calendar Authentication")
        title_font = QFont()
        title_font.setPointSize(12)
        title_font.setBold(True)
        title.setFont(title_font)
        layout.addWidget(title)
        
        # Status box
        status_box = QGroupBox("Authentication Status")
        status_layout = QVBoxLayout(status_box)
        self.auth_status_label = QLabel("Checking authentication...")
        status_layout.addWidget(self.auth_status_label)
        layout.addWidget(status_box)

        calendar_box = QGroupBox("Calendar Selection")
        calendar_layout = QFormLayout(calendar_box)
        self.calendar_combo = QComboBox()
        self.calendar_combo.currentIndexChanged.connect(self.on_calendar_selection_changed)
        refresh_calendars_btn = QPushButton("Refresh Calendars")
        refresh_calendars_btn.clicked.connect(self.populate_calendars)
        calendar_layout.addRow("Sync Target Calendar:", self.calendar_combo)
        calendar_layout.addRow("", refresh_calendars_btn)
        layout.addWidget(calendar_box)
        
        # Action buttons
        button_layout = QHBoxLayout()
        
        auth_btn = QPushButton("Authenticate")
        auth_btn.clicked.connect(self.authenticate_google)
        auth_btn.setStyleSheet("background-color: #2196F3; color: white; padding: 10px;")
        button_layout.addWidget(auth_btn)
        
        logout_btn = QPushButton("Logout")
        logout_btn.clicked.connect(self.logout_google)
        logout_btn.setStyleSheet("background-color: #f44336; color: white; padding: 10px;")
        button_layout.addWidget(logout_btn)
        
        button_layout.addStretch()
        layout.addLayout(button_layout)
        
        # Info text
        info_text = QTextEdit()
        info_text.setReadOnly(True)
        info_text.setText(
            "Click 'Authenticate' to authorize the sync agent to access your Google Calendar.\n\n"
            "A browser window will open asking for your Google account credentials.\n"
            "Your authorization token is stored locally in token.json.\n\n"
            "Click 'Logout' to revoke the current authorization."
        )
        layout.addWidget(info_text)
        
        layout.addStretch()
        self.update_auth_status()
        self.populate_calendars()
        return widget
    
    def create_history_tab(self) -> QWidget:
        """Create the sync history tab."""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # Title
        title = QLabel("Recent Sync History")
        title_font = QFont()
        title_font.setPointSize(12)
        title_font.setBold(True)
        title.setFont(title_font)
        layout.addWidget(title)
        
        # History box
        history_box = QGroupBox("Last Sync")
        history_layout = QFormLayout(history_box)
        
        self.last_sync_time_label = QLabel("Never")
        self.last_sync_status_label = QLabel("Never run")
        self.last_sync_events_label = QLabel("")
        
        history_layout.addRow("Last Sync Time:", self.last_sync_time_label)
        history_layout.addRow("Status:", self.last_sync_status_label)
        history_layout.addRow("Events:", self.last_sync_events_label)
        
        layout.addWidget(history_box)
        
        # Log viewer
        log_box = QGroupBox("Sync Log")
        log_layout = QVBoxLayout(log_box)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setMaximumHeight(300)
        log_layout.addWidget(self.log_text)
        
        log_actions_layout = QHBoxLayout()

        refresh_log_btn = QPushButton("Refresh Log")
        refresh_log_btn.clicked.connect(self.refresh_log)
        log_actions_layout.addWidget(refresh_log_btn)

        clear_log_btn = QPushButton("Clear Log")
        clear_log_btn.clicked.connect(self.clear_log)
        log_actions_layout.addWidget(clear_log_btn)

        sync_now_btn = QPushButton("Sync Now")
        sync_now_btn.clicked.connect(self.sync_now)
        sync_now_btn.setStyleSheet("background-color: #2e7d32; color: white; padding: 6px 10px; font-weight: bold;")
        log_actions_layout.addWidget(sync_now_btn)

        log_actions_layout.addStretch()
        log_layout.addLayout(log_actions_layout)
        
        layout.addWidget(log_box)
        
        layout.addStretch()
        self.refresh_history()
        return widget
    
    def load_settings(self):
        """Load settings from config and populate UI."""
        self.outlook_path_input.setText(self.config_manager.get_outlook_json_path())
        self.timezone_combo.setCurrentText(self.config_manager.get_timezone())
        self.sync_frequency_spin.setValue(self.config_manager.get_sync_frequency_minutes())
        self.logging_level_combo.setCurrentText(self.config_manager.get_logging_level())
        self.selected_calendar_label.setText(
            f"{self.config_manager.get_google_calendar_name()} ({self.config_manager.get_google_calendar_id()})"
        )
    
    def save_settings(self):
        """(Deprecated) Save settings from UI to config. No longer needed with autosave."""
        pass

    def autosave_setting(self, key, value):
        """Autosave a single setting change."""
        try:
            self.config_manager.set(key, value)
            # Update label if calendar name/id
            if key in ("GOOGLE_CALENDAR_ID", "GOOGLE_CALENDAR_NAME"):
                self.selected_calendar_label.setText(
                    f"{self.config_manager.get_google_calendar_name()} ({self.config_manager.get_google_calendar_id()})"
                )
        except Exception as e:
            QMessageBox.critical(self, "Autosave Error", f"Failed to autosave {key}: {e}")
    
    def browse_outlook_file(self):
        """Open file browser for Outlook snapshot file."""
        file_path, _ = QFileDialog.getOpenFileName(
            self,
            "Select Outlook Snapshot JSON",
            "",
            "JSON Files (*.json);;All Files (*)"
        )
        if file_path:
            self.outlook_path_input.setText(file_path)
    
    def authenticate_google(self):
        """Authenticate with Google Calendar."""
        try:
            self.status_label.setText("Status: Authenticating with Google...")
            QApplication.processEvents()
            
            get_google_service(self.config_manager.get_timezone())
            QMessageBox.information(self, "Success", "Successfully authenticated with Google Calendar!")
            self.update_auth_status()
            self.populate_calendars()
            self.status_label.setText("Status: Ready")
            
        except Exception as e:
            QMessageBox.critical(self, "Authentication Error", f"Failed to authenticate: {e}")
            self.status_label.setText("Status: Authentication failed")
    
    def logout_google(self):
        """Logout by deleting the token file."""
        try:
            if os.path.exists('token.json'):
                os.remove('token.json')
                self.config_manager.set_google_calendar('primary', 'Primary Calendar')
                self.calendar_combo.clear()
                self.calendar_combo.addItem('Primary Calendar', 'primary')
                QMessageBox.information(self, "Success", "Successfully logged out!")
                self.update_auth_status()
                self.selected_calendar_label.setText('Primary Calendar (primary)')
            else:
                QMessageBox.information(self, "Info", "No active authentication found.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to logout: {e}")
    
    def update_auth_status(self):
        """Update the authentication status display."""
        if os.path.exists('token.json'):
            try:
                service = get_google_service(self.config_manager.get_timezone())
                calendars = list_google_calendars(service)
                primary = next((c for c in calendars if c.get('primary')), None)
                account_hint = primary['id'] if primary else self.config_manager.get_google_calendar_id()
                self.auth_status_label.setText(f"Authenticated {account_hint}")
                self.auth_status_label.setStyleSheet("color: green; font-weight: bold;")
            except Exception:
                self.auth_status_label.setText("Authenticated (token present). Unable to resolve account hint.")
                self.auth_status_label.setStyleSheet("color: #b36b00; font-weight: bold;")
        else:
            self.auth_status_label.setText("Not authenticated")
            self.auth_status_label.setStyleSheet("color: red; font-weight: bold;")

    def populate_calendars(self):
        """Load available Google calendars and populate selector."""
        self.calendar_combo.clear()
        if not os.path.exists('token.json'):
            self.calendar_combo.addItem('Authenticate to load calendars', '')
            return

        try:
            service = get_google_service(self.config_manager.get_timezone())
            calendars = list_google_calendars(service)
            if not calendars:
                self.calendar_combo.addItem('No calendars found', '')
                return

            # Primary calendars first, then by summary.
            calendars.sort(key=lambda c: (not c.get('primary', False), c.get('summary', '').lower()))
            for cal in calendars:
                label = cal['summary']
                if cal.get('primary'):
                    label = f"{label} [Primary]"
                self.calendar_combo.addItem(label, cal['id'])

            saved_id = self.config_manager.get_google_calendar_id()
            idx = self.calendar_combo.findData(saved_id)
            if idx >= 0:
                self.calendar_combo.setCurrentIndex(idx)
            else:
                self.calendar_combo.setCurrentIndex(0)
                self.on_calendar_selection_changed(0)
        except Exception as e:
            self.calendar_combo.clear()
            self.calendar_combo.addItem(f"Failed to load calendars: {e}", '')

    def on_calendar_selection_changed(self, index: int):
        """Persist selected calendar and update display label (autosave)."""
        calendar_id = self.calendar_combo.itemData(index)
        calendar_name = self.calendar_combo.itemText(index)
        if not calendar_id:
            return
        self.autosave_setting('GOOGLE_CALENDAR_ID', calendar_id)
        self.autosave_setting('GOOGLE_CALENDAR_NAME', calendar_name)
        self.selected_calendar_label.setText(f"{calendar_name} ({calendar_id})")
    
    def refresh_history(self):
        """Refresh the sync history display."""
        last_sync = self.config_manager.get_last_sync_info()
        
        if last_sync['time']:
            try:
                parsed = datetime.fromisoformat(str(last_sync['time']))
                self.last_sync_time_label.setText(parsed.strftime('%Y-%m-%d %H:%M:%S'))
            except ValueError:
                # Backward-compatible fallback for older timestamp formats.
                self.last_sync_time_label.setText(str(last_sync['time']).split('.')[0])
        else:
            self.last_sync_time_label.setText("Never")
        
        status = last_sync['status']
        if status.startswith('Success'):
            self.last_sync_status_label.setText(status)
            self.last_sync_status_label.setStyleSheet("color: green;")
        elif status.startswith('Error'):
            self.last_sync_status_label.setText(status)
            self.last_sync_status_label.setStyleSheet("color: red;")
        else:
            self.last_sync_status_label.setText(status)
            self.last_sync_status_label.setStyleSheet("color: orange;")
        
        events_text = (
            f"Created: {last_sync['created']}, "
            f"Updated: {last_sync['updated']}, "
            f"Deleted: {last_sync['deleted']}"
        )
        self.last_sync_events_label.setText(events_text)
        
        self.refresh_log()
    
    def refresh_log(self):
        """Refresh the log file display."""
        try:
            if os.path.exists('sync.log'):
                with open('sync.log', 'r', encoding='utf-8') as f:
                    lines = f.readlines()
                    # Show last 50 lines
                    recent_lines = lines[-50:] if len(lines) > 50 else lines
                    self.log_text.setText(''.join(recent_lines))
                    self.log_text.verticalScrollBar().setValue(
                        self.log_text.verticalScrollBar().maximum()
                    )
            else:
                self.log_text.setText("No log file found yet.")
        except Exception as e:
            self.log_text.setText(f"Error reading log: {e}")

    def clear_log(self):
        """Clear all entries from the sync log file and viewer."""
        try:
            with open('sync.log', 'w', encoding='utf-8'):
                pass
            self.log_text.clear()
            self.status_label.setText("Status: Log cleared")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to clear log: {e}")

    def sync_now(self):
        """Run a sync immediately from the UI."""
        try:
            outlook_json_path = self.config_manager.get_outlook_json_path()
            if not outlook_json_path:
                QMessageBox.warning(self, "Missing Configuration", "Please select an Outlook snapshot file first.")
                return
            if not os.path.exists(outlook_json_path):
                QMessageBox.warning(self, "File Not Found", f"Cannot find snapshot file:\n{outlook_json_path}")
                return

            timezone = self.config_manager.get_timezone()
            calendar_id = self.config_manager.get_google_calendar_id()

            self.status_label.setText("Status: Syncing now...")
            QApplication.processEvents()

            service = get_google_service(timezone)
            outlook_events = load_outlook_snapshot(outlook_json_path)
            stats = perform_sync(service, outlook_events, timezone, calendar_id)

            self.config_manager.update_sync_result(
                'Success',
                stats.get('created', 0),
                stats.get('updated', 0),
                stats.get('deleted', 0),
            )

            self.status_label.setText(
                f"Status: Sync complete (C:{stats.get('created', 0)} U:{stats.get('updated', 0)} D:{stats.get('deleted', 0)})"
            )
            self.refresh_history()
        except Exception as e:
            self.config_manager.update_sync_result(f'Error: {str(e)}')
            self.status_label.setText("Status: Sync failed")
            self.refresh_history()
            QMessageBox.critical(self, "Sync Error", f"Sync failed: {e}")
    
    def refresh_status(self):
        """Periodically refresh status information."""
        self.refresh_history()

    def setup_system_tray(self):
        """Create system tray icon and actions for close-to-tray behavior."""
        from PyQt5.QtGui import QIcon
        if not QSystemTrayIcon.isSystemTrayAvailable():
            return

        tray_menu = QMenu(self)
        show_action = QAction("Open Settings", self)
        quit_action = QAction("Quit", self)
        show_action.triggered.connect(self.restore_from_tray)
        quit_action.triggered.connect(self.quit_from_tray)
        tray_menu.addAction(show_action)
        tray_menu.addAction(quit_action)

        self.tray_icon = QSystemTrayIcon(self)
        # Use logo as tray icon if available
        if os.path.exists(self.logo_path):
            self.tray_icon.setIcon(QIcon(self.logo_path))
        else:
            self.tray_icon.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))
        self.tray_icon.setToolTip("Outlook to Google Sync")
        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.activated.connect(self.on_tray_activated)
        self.tray_icon.show()

    def on_tray_activated(self, reason):
        """Restore the window on tray icon activation."""
        if reason == QSystemTrayIcon.Trigger:
            self.restore_from_tray()

    def restore_from_tray(self):
        """Restore and focus the settings window."""
        self.showNormal()
        self.activateWindow()
        self.raise_()

    def quit_from_tray(self):
        """Quit the application from tray menu."""
        self.is_quitting = True
        if self.tray_icon:
            self.tray_icon.hide()
        QApplication.instance().quit()

    def closeEvent(self, event):
        """Minimize to tray on close so monitoring can continue."""
        if self.tray_icon and self.tray_icon.isVisible() and not self.is_quitting:
            event.ignore()
            self.hide()
            if not self.has_shown_tray_hint:
                self.tray_icon.showMessage(
                    "Outlook to Google Sync",
                    "App minimized to tray. Use tray icon to reopen or quit.",
                    QSystemTrayIcon.Information,
                    4000,
                )
                self.has_shown_tray_hint = True
        else:
            event.accept()


def main():
    """Main entry point for the GUI."""
    app = QApplication(sys.argv)
    window = SyncSettingsApp()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
