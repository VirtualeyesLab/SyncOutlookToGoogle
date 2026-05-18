"""
PyQt5 GUI for Outlook to Google Calendar Sync Agent.
Provides UI for managing settings, authentication, and viewing sync history.
"""

import sys
import os
import json
from pathlib import Path
from datetime import datetime
from typing import Optional

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QTabWidget, QLabel, QLineEdit, QPushButton, QSpinBox, QComboBox,
    QFileDialog, QTextEdit, QStatusBar, QMessageBox, QGridLayout,
    QGroupBox, QFormLayout
)
from PyQt5.QtCore import Qt, QTimer, pyqtSignal, QObject
from PyQt5.QtGui import QFont, QColor

from agent_config import ConfigManager
from sync import get_google_service


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
        self.init_ui()
        self.load_settings()
        
        # Timer to refresh status
        self.status_timer = QTimer()
        self.status_timer.timeout.connect(self.refresh_status)
        self.status_timer.start(5000)  # Update every 5 seconds
    
    def init_ui(self):
        """Initialize the user interface."""
        self.setWindowTitle("Outlook to Google Sync - Settings")
        self.setGeometry(100, 100, 900, 700)
        
        # Create central widget and main layout
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        
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
        
        # Monitoring enabled
        monitoring_layout = QHBoxLayout()
        self.monitoring_check_label = QLabel("✓ File monitoring enabled")
        self.monitoring_check_label.setStyleSheet("color: green; font-weight: bold;")
        monitoring_layout.addWidget(self.monitoring_check_label)
        monitoring_layout.addStretch()
        form_layout.addRow("Monitoring:", monitoring_layout)
        
        layout.addLayout(form_layout)
        
        # Save button
        save_btn = QPushButton("Save Settings")
        save_btn.clicked.connect(self.save_settings)
        save_btn.setStyleSheet("background-color: #4CAF50; color: white; padding: 10px; font-weight: bold;")
        layout.addWidget(save_btn)
        
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
        
        refresh_log_btn = QPushButton("Refresh Log")
        refresh_log_btn.clicked.connect(self.refresh_log)
        log_layout.addWidget(refresh_log_btn)
        
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
    
    def save_settings(self):
        """Save settings from UI to config."""
        try:
            self.config_manager.update({
                'OUTLOOK_JSON_PATH': self.outlook_path_input.text(),
                'TIMEZONE': self.timezone_combo.currentText(),
                'SYNC_FREQUENCY_MINUTES': self.sync_frequency_spin.value(),
                'LOGGING_LEVEL': self.logging_level_combo.currentText(),
            })
            QMessageBox.information(self, "Success", "Settings saved successfully!")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to save settings: {e}")
    
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
            
            service = get_google_service(self.config_manager.get_timezone())
            QMessageBox.information(self, "Success", "Successfully authenticated with Google Calendar!")
            self.update_auth_status()
            self.status_label.setText("Status: Ready")
            
        except Exception as e:
            QMessageBox.critical(self, "Authentication Error", f"Failed to authenticate: {e}")
            self.status_label.setText("Status: Authentication failed")
    
    def logout_google(self):
        """Logout by deleting the token file."""
        try:
            if os.path.exists('token.json'):
                os.remove('token.json')
                QMessageBox.information(self, "Success", "Successfully logged out!")
                self.update_auth_status()
            else:
                QMessageBox.information(self, "Info", "No active authentication found.")
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to logout: {e}")
    
    def update_auth_status(self):
        """Update the authentication status display."""
        if os.path.exists('token.json'):
            try:
                with open('token.json', 'r') as f:
                    token_data = json.load(f)
                    email = token_data.get('email', 'Unknown')
                self.auth_status_label.setText(
                    f"✓ Authenticated as: {email}"
                )
                self.auth_status_label.setStyleSheet("color: green; font-weight: bold;")
            except:
                self.auth_status_label.setText("✗ Token file corrupted")
                self.auth_status_label.setStyleSheet("color: red; font-weight: bold;")
        else:
            self.auth_status_label.setText("✗ Not authenticated")
            self.auth_status_label.setStyleSheet("color: red; font-weight: bold;")
    
    def refresh_history(self):
        """Refresh the sync history display."""
        last_sync = self.config_manager.get_last_sync_info()
        
        if last_sync['time']:
            self.last_sync_time_label.setText(last_sync['time'])
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
    
    def refresh_status(self):
        """Periodically refresh status information."""
        self.update_auth_status()
        self.refresh_history()


def main():
    """Main entry point for the GUI."""
    app = QApplication(sys.argv)
    window = SyncSettingsApp()
    window.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()
