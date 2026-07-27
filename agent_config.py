"""
Configuration management for the Outlook to Google sync agent.
Handles extended configuration schema with settings for timezone, 
sync frequency, monitoring, and sync history.
"""

import json
import os
from datetime import datetime
from typing import Optional, Dict, Any


class ConfigManager:
    """
    Manages persistent configuration for the sync agent.
    Handles loading, saving, and schema migration.
    """
    
    # Default configuration schema
    DEFAULT_CONFIG = {
        'OUTLOOK_JSON_PATH': '',
        'TIMEZONE': 'Pacific/Auckland',
        'GOOGLE_CALENDAR_ID': 'primary',
        'GOOGLE_CALENDAR_NAME': 'Primary Calendar',
        'SYNC_FREQUENCY_MINUTES': 15,
        'MONITORING_ENABLED': True,
        'LOGGING_LEVEL': 'INFO',
        'LAST_SYNC_TIME': None,
        'LAST_SYNC_STATUS': 'Never run',
        'LAST_SYNC_CREATED': 0,
        'LAST_SYNC_UPDATED': 0,
        'LAST_SYNC_DELETED': 0,
    }
    
    def __init__(self, config_file: str = 'config.json'):
        self.config_file = config_file
        self.config = self.load()
    
    def load(self) -> Dict[str, Any]:
        """Load configuration from file, migrating schema if needed."""
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    loaded = json.load(f)
                
                # Migrate old config format to new schema
                if 'OUTLOOK_JSON_PATH' in loaded and len(loaded) == 1:
                    # Old format: only OUTLOOK_JSON_PATH
                    loaded = self._migrate_old_config(loaded)
                
                # Ensure all keys from DEFAULT_CONFIG exist
                config = self.DEFAULT_CONFIG.copy()
                config.update(loaded)
                return config
                
            except (json.JSONDecodeError, IOError) as e:
                print(f"Warning: Could not load config file: {e}")
                return self.DEFAULT_CONFIG.copy()
        
        return self.DEFAULT_CONFIG.copy()
    
    def _migrate_old_config(self, old_config: Dict[str, Any]) -> Dict[str, Any]:
        """Migrate old config format to new schema."""
        new_config = self.DEFAULT_CONFIG.copy()
        new_config['OUTLOOK_JSON_PATH'] = old_config.get('OUTLOOK_JSON_PATH', '')
        print(f"Migrated config from old format to new schema.")
        return new_config
    
    def save(self):
        """Save current configuration to file."""
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4, default=str)
        except IOError as e:
            print(f"Error saving config: {e}")
    
    def get(self, key: str, default: Any = None) -> Any:
        """Get a configuration value."""
        return self.config.get(key, default)
    
    def set(self, key: str, value: Any):
        """Set a configuration value and save."""
        self.config[key] = value
        self.save()
    
    def update(self, updates: Dict[str, Any]):
        """Update multiple configuration values and save."""
        self.config.update(updates)
        self.save()
    
    def update_sync_result(self, status: str, created: int = 0, updated: int = 0, deleted: int = 0):
        """Update sync result statistics."""
        self.update({
            'LAST_SYNC_TIME': datetime.now().isoformat(timespec='seconds'),
            'LAST_SYNC_STATUS': status,
            'LAST_SYNC_CREATED': created,
            'LAST_SYNC_UPDATED': updated,
            'LAST_SYNC_DELETED': deleted,
        })
    
    def get_outlook_json_path(self) -> str:
        """Get the configured Outlook event JSON folder path."""
        return self.get('OUTLOOK_JSON_PATH', '')
    
    def set_outlook_json_path(self, path: str):
        """Set the Outlook event JSON folder path."""
        self.set('OUTLOOK_JSON_PATH', path)
    
    def get_timezone(self) -> str:
        """Get the configured timezone."""
        return self.get('TIMEZONE', 'Pacific/Auckland')
    
    def set_timezone(self, tz: str):
        """Set the timezone."""
        self.set('TIMEZONE', tz)
    
    def get_sync_frequency_minutes(self) -> int:
        """Get sync frequency in minutes."""
        return self.get('SYNC_FREQUENCY_MINUTES', 15)

    def get_google_calendar_id(self) -> str:
        """Get target Google Calendar ID."""
        return self.get('GOOGLE_CALENDAR_ID', 'primary')

    def set_google_calendar(self, calendar_id: str, calendar_name: str = ''):
        """Set target Google Calendar ID and display name."""
        updates = {'GOOGLE_CALENDAR_ID': calendar_id or 'primary'}
        if calendar_name:
            updates['GOOGLE_CALENDAR_NAME'] = calendar_name
        self.update(updates)

    def get_google_calendar_name(self) -> str:
        """Get target Google Calendar display name."""
        return self.get('GOOGLE_CALENDAR_NAME', 'Primary Calendar')
    
    def set_sync_frequency_minutes(self, minutes: int):
        """Set sync frequency in minutes."""
        self.set('SYNC_FREQUENCY_MINUTES', max(1, minutes))
    
    def is_monitoring_enabled(self) -> bool:
        """Check if file monitoring is enabled."""
        return self.get('MONITORING_ENABLED', True)
    
    def set_monitoring_enabled(self, enabled: bool):
        """Enable or disable file monitoring."""
        self.set('MONITORING_ENABLED', enabled)
    
    def get_logging_level(self) -> str:
        """Get the logging level."""
        return self.get('LOGGING_LEVEL', 'INFO')
    
    def set_logging_level(self, level: str):
        """Set the logging level."""
        valid_levels = ['DEBUG', 'INFO', 'WARNING', 'ERROR', 'CRITICAL']
        if level.upper() in valid_levels:
            self.set('LOGGING_LEVEL', level.upper())
    
    def get_last_sync_info(self) -> Dict[str, Any]:
        """Get information about the last sync."""
        return {
            'time': self.get('LAST_SYNC_TIME'),
            'status': self.get('LAST_SYNC_STATUS'),
            'created': self.get('LAST_SYNC_CREATED'),
            'updated': self.get('LAST_SYNC_UPDATED'),
            'deleted': self.get('LAST_SYNC_DELETED'),
        }
