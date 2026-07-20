"""
Sync Agent: Monitors OutlookSnapshot.json for changes and automatically triggers syncs.
Supports both watchdog-based file system monitoring and polling fallback.
"""

import os
import sys
import time
import logging
import threading
import argparse
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional, Callable

try:
    from watchdog.observers import Observer
    from watchdog.events import FileSystemEventHandler
    WATCHDOG_AVAILABLE = True
except ImportError:
    WATCHDOG_AVAILABLE = False
    Observer = None
    FileSystemEventHandler = None
    logging.warning("watchdog not installed. File monitoring will use polling fallback.")

from sync import get_google_service, load_outlook_snapshot, perform_sync
from agent_config import ConfigManager


# Configure logging
def setup_logging(level: str = 'INFO'):
    """Configure logging for the agent."""
    log_level = getattr(logging, level.upper(), logging.INFO)
    logging.basicConfig(
        level=log_level,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler("sync.log", encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    logging.getLogger("googleapiclient.discovery_cache").setLevel(logging.WARNING)


logger = logging.getLogger(__name__)


# Define OutlookFileHandler only if watchdog is available
if WATCHDOG_AVAILABLE:
    class OutlookFileHandler(FileSystemEventHandler):
        """Watchdog file system event handler for OutlookSnapshot.json changes."""
        
        def __init__(self, callback: Callable, debounce_seconds: float = 2.0):
            super().__init__()
            self.callback = callback
            self.debounce_seconds = debounce_seconds
            self.last_event_time = 0
        
        def on_modified(self, event):
            """Called when a file is modified."""
            if event.is_directory:
                return
            
            # Debounce: ignore events closer than debounce_seconds apart
            current_time = time.time()
            if current_time - self.last_event_time < self.debounce_seconds:
                return
            
            self.last_event_time = current_time
            self.callback()
else:
    class OutlookFileHandler:
        """Placeholder when watchdog is not available."""
        def __init__(self, callback: Callable, debounce_seconds: float = 2.0):
            pass
        
        def on_modified(self, event):
            pass


class SyncAgent:
    """
    Monitors the Outlook snapshot file and triggers syncs when changes are detected.
    Supports both watchdog (real-time) and polling (fallback) monitoring.
    """
    
    def __init__(self, config_file: str = 'config.json', use_polling: bool = False):
        self.config_manager = ConfigManager(config_file)
        self.config_file = config_file
        self.use_polling = use_polling or not WATCHDOG_AVAILABLE
        self.running = False
        self.observer = None
        self.polling_thread = None
        self.last_sync_time = 0
        self.min_sync_interval = 60  # Prevent syncs faster than 60 seconds
        
        # Setup logging
        setup_logging(self.config_manager.get_logging_level())
        
        logger.info("SyncAgent initialized")
        logger.info(f"Using {'polling' if self.use_polling else 'watchdog'} for file monitoring")
    
    def _get_outlook_snapshot_path(self) -> Optional[str]:
        """Get the path to the Outlook snapshot file."""
        path = self.config_manager.get_outlook_json_path()
        if not path:
            logger.error("Outlook snapshot path not configured. Run sync.py first.")
            return None
        return path
    
    def _can_sync(self) -> bool:
        """Check if enough time has passed since last sync."""
        current_time = time.time()
        if current_time - self.last_sync_time < self.min_sync_interval:
            return False
        return True
    
    def _perform_sync_safe(self):
        """Perform sync with error handling."""
        if not self._can_sync():
            logger.debug("Sync rate-limited, skipping")
            return
        
        snapshot_path = self._get_outlook_snapshot_path()
        if not snapshot_path or not os.path.exists(snapshot_path):
            logger.error(f"Snapshot file not found: {snapshot_path}")
            self.config_manager.update_sync_result('Error: Snapshot file not found')
            return
        
        try:
            logger.info("Snapshot file changed, triggering sync...")
            
            # Load snapshot and get service
            outlook_events = load_outlook_snapshot(snapshot_path)
            timezone = self.config_manager.get_timezone()
            calendar_id = self.config_manager.get_google_calendar_id()
            service = get_google_service(timezone)
            
            # Perform sync
            stats = perform_sync(service, outlook_events, timezone, calendar_id)
            
            # Update sync results in config
            self.config_manager.update_sync_result('Success', stats['created'], stats['updated'], stats['deleted'])
            self.last_sync_time = time.time()
            
            logger.info(f"Sync complete: Created {stats['created']}, Updated {stats['updated']}, Deleted {stats['deleted']}")
            
        except Exception as e:
            logger.error(f"Sync failed: {e}", exc_info=True)
            self.config_manager.update_sync_result(f'Error: {str(e)}')
    
    def _start_watchdog_monitoring(self, snapshot_path: str):
        """Start watchdog-based file monitoring."""
        if not WATCHDOG_AVAILABLE:
            logger.warning("Watchdog not available, falling back to polling")
            self.use_polling = True
            self._start_polling_monitoring(snapshot_path)
            return
        
        try:
            snapshot_dir = os.path.dirname(snapshot_path) or '.'
            snapshot_name = os.path.basename(snapshot_path)
            
            event_handler = OutlookFileHandler(self._perform_sync_safe, debounce_seconds=2.0)
            self.observer = Observer()
            self.observer.schedule(event_handler, snapshot_dir, recursive=False)
            self.observer.start()
            
            logger.info(f"Watching {snapshot_dir} for changes to {snapshot_name}")
            
            # Keep observer running
            while self.running:
                time.sleep(1)
            
            self.observer.stop()
            self.observer.join()
            
        except Exception as e:
            logger.error(f"Watchdog monitoring failed: {e}")
            logger.info("Falling back to polling")
            self.use_polling = True
            self._start_polling_monitoring(snapshot_path)
    
    def _start_polling_monitoring(self, snapshot_path: str):
        """Start polling-based file monitoring (fallback)."""
        poll_interval = max(5, self.config_manager.get_sync_frequency_minutes() * 60)
        last_mtime = 0
        
        logger.info(f"Polling {snapshot_path} every {poll_interval} seconds")
        
        while self.running:
            try:
                if os.path.exists(snapshot_path):
                    current_mtime = os.path.getmtime(snapshot_path)
                    if current_mtime != last_mtime:
                        last_mtime = current_mtime
                        self._perform_sync_safe()
                
                time.sleep(poll_interval)
                
            except Exception as e:
                logger.error(f"Polling monitoring error: {e}")
                time.sleep(poll_interval)
    
    def start(self):
        """Start the monitoring agent."""
        if self.running:
            logger.warning("Agent already running")
            return
        
        self.running = True
        snapshot_path = self._get_outlook_snapshot_path()
        
        if not snapshot_path:
            logger.error("Cannot start agent: no snapshot path configured")
            self.running = False
            return
        
        if not os.path.exists(snapshot_path):
            logger.warning(f"Snapshot file does not exist yet: {snapshot_path}")
            logger.info("Agent will monitor for file creation...")
        
        logger.info("Starting Outlook-to-Google sync agent...")
        
        # Start monitoring in main thread (blocking)
        if not self.use_polling and WATCHDOG_AVAILABLE:
            self._start_watchdog_monitoring(snapshot_path)
        else:
            self._start_polling_monitoring(snapshot_path)
    
    def stop(self):
        """Stop the monitoring agent."""
        logger.info("Stopping agent...")
        self.running = False
        
        if self.observer:
            self.observer.stop()
            self.observer.join()
        
        logger.info("Agent stopped")


def main():
    """CLI entry point for the agent."""
    parser = argparse.ArgumentParser(
        description='Outlook to Google Calendar sync agent with file monitoring'
    )
    parser.add_argument(
        '--monitor',
        action='store_true',
        help='Start the monitoring agent'
    )
    parser.add_argument(
        '--sync-now',
        action='store_true',
        help='Perform a single sync and exit'
    )
    parser.add_argument(
        '--polling-only',
        action='store_true',
        help='Use polling instead of watchdog (for debugging)'
    )
    
    args = parser.parse_args()
    
    if args.sync_now:
        # Single sync
        logger.info("Performing single sync...")
        from sync import main as sync_main
        sync_main()
    elif args.monitor:
        # Start monitoring agent
        agent = SyncAgent(use_polling=args.polling_only)
        try:
            agent.start()
        except KeyboardInterrupt:
            logger.info("Received interrupt, shutting down...")
            agent.stop()
    else:
        parser.print_help()


if __name__ == '__main__':
    main()
