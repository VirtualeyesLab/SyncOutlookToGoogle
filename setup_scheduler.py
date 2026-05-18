"""
Windows Task Scheduler setup for the Outlook to Google sync agent.
Creates a scheduled task that runs the agent at startup.
"""

import os
import sys
import subprocess
import json
from pathlib import Path
from typing import Optional


class TaskSchedulerManager:
    """Manage Windows Task Scheduler tasks for the sync agent."""
    
    TASK_NAME = "SyncOutlookToGoogle"
    TASK_DESCRIPTION = "Monitors Outlook calendar and syncs events to Google Calendar"
    
    @staticmethod
    def get_python_executable() -> str:
        """Get the path to the current Python executable."""
        return sys.executable
    
    @staticmethod
    def get_script_path() -> str:
        """Get the path to the agent.py script."""
        return os.path.join(os.path.dirname(__file__), 'agent.py')
    
    @classmethod
    def create_task(cls, run_at_startup: bool = True) -> bool:
        """
        Create a Windows Task Scheduler task to run the agent.
        
        Args:
            run_at_startup: If True, task runs at user logon; if False, runs every 15 minutes
        
        Returns:
            True if successful, False otherwise
        """
        try:
            python_exe = cls.get_python_executable()
            script_path = cls.get_script_path()
            
            if not os.path.exists(script_path):
                print(f"Error: Script not found at {script_path}")
                return False
            
            # Build the command
            command = f'"{python_exe}" "{script_path}" --monitor'
            
            # Create the task using schtasks command
            if run_at_startup:
                # Run at user logon
                cmd = [
                    'schtasks', '/create',
                    '/tn', cls.TASK_NAME,
                    '/tr', command,
                    '/sc', 'onlogon',
                    '/rl', 'highest',
                    '/f',
                    '/it'  # Run only if user is logged in
                ]
                trigger_desc = "At user logon"
            else:
                # Run every 15 minutes
                cmd = [
                    'schtasks', '/create',
                    '/tn', cls.TASK_NAME,
                    '/tr', command,
                    '/sc', 'minute',
                    '/mo', '15',
                    '/rl', 'highest',
                    '/f'
                ]
                trigger_desc = "Every 15 minutes"
            
            print(f"Creating Task Scheduler task: {cls.TASK_NAME}")
            print(f"Trigger: {trigger_desc}")
            print(f"Command: {command}")
            
            result = subprocess.run(cmd, capture_output=True, text=True)
            
            if result.returncode == 0:
                print("✓ Task created successfully")
                return True
            else:
                print(f"✗ Failed to create task: {result.stderr}")
                return False
        
        except Exception as e:
            print(f"Error creating task: {e}")
            return False
    
    @classmethod
    def delete_task(cls) -> bool:
        """
        Delete the Windows Task Scheduler task.
        
        Returns:
            True if successful, False otherwise
        """
        try:
            cmd = ['schtasks', '/delete', '/tn', cls.TASK_NAME, '/f']
            
            print(f"Deleting Task Scheduler task: {cls.TASK_NAME}")
            result = subprocess.run(cmd, capture_output=True, text=True)
            
            if result.returncode == 0:
                print("✓ Task deleted successfully")
                return True
            else:
                print(f"✗ Failed to delete task: {result.stderr}")
                return False
        
        except Exception as e:
            print(f"Error deleting task: {e}")
            return False
    
    @classmethod
    def task_exists(cls) -> bool:
        """Check if the task already exists."""
        try:
            cmd = ['schtasks', '/query', '/tn', cls.TASK_NAME]
            result = subprocess.run(cmd, capture_output=True, text=True)
            return result.returncode == 0
        except:
            return False
    
    @classmethod
    def query_task(cls) -> Optional[str]:
        """Query task details."""
        try:
            cmd = ['schtasks', '/query', '/tn', cls.TASK_NAME, '/v', '/fo', 'list']
            result = subprocess.run(cmd, capture_output=True, text=True)
            if result.returncode == 0:
                return result.stdout
            return None
        except:
            return None


def check_admin():
    """Check if running with administrator privileges."""
    try:
        # On Windows, check if running with admin privileges
        import ctypes
        return ctypes.windll.shell.IsUserAnAdmin()
    except:
        return False


def main():
    """Main entry point for task scheduler setup."""
    import argparse
    
    parser = argparse.ArgumentParser(
        description='Setup Windows Task Scheduler for Outlook to Google sync agent'
    )
    parser.add_argument(
        '--install',
        action='store_true',
        help='Create the scheduled task'
    )
    parser.add_argument(
        '--uninstall',
        action='store_true',
        help='Delete the scheduled task'
    )
    parser.add_argument(
        '--status',
        action='store_true',
        help='Show task status'
    )
    parser.add_argument(
        '--frequency',
        choices=['startup', 'interval'],
        default='startup',
        help='Set when the agent runs (default: startup)'
    )
    
    args = parser.parse_args()
    
    # Check for admin privileges
    if not check_admin():
        print("Error: This script requires administrator privileges.")
        print("Please run as Administrator (right-click cmd.exe > Run as administrator)")
        sys.exit(1)
    
    if args.install:
        run_at_startup = args.frequency == 'startup'
        success = TaskSchedulerManager.create_task(run_at_startup=run_at_startup)
        sys.exit(0 if success else 1)
    
    elif args.uninstall:
        success = TaskSchedulerManager.delete_task()
        sys.exit(0 if success else 1)
    
    elif args.status:
        if TaskSchedulerManager.task_exists():
            print(f"✓ Task '{TaskSchedulerManager.TASK_NAME}' exists")
            details = TaskSchedulerManager.query_task()
            if details:
                print("\nTask details:")
                print(details)
        else:
            print(f"✗ Task '{TaskSchedulerManager.TASK_NAME}' does not exist")
        sys.exit(0)
    
    else:
        # Default: check status and offer to install
        if TaskSchedulerManager.task_exists():
            print(f"✓ Task '{TaskSchedulerManager.TASK_NAME}' already exists")
            print("\nTo uninstall, run: python setup_scheduler.py --uninstall")
        else:
            print(f"Task '{TaskSchedulerManager.TASK_NAME}' does not exist")
            print("\nTo install, run: python setup_scheduler.py --install")
            print("(Administrator privileges required)")


if __name__ == '__main__':
    main()
