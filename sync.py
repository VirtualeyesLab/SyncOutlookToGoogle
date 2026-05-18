import datetime
import json
import os
import sys
import logging
from typing import Dict, List, Tuple, Optional
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from agent_config import ConfigManager

# Configure Logging (Writes to terminal AND sync.log file)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("sync.log", encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

SCOPES = ['https://www.googleapis.com/auth/calendar']
CONFIG_FILE = 'config.json'


def get_config() -> str:
    """Get the Outlook snapshot file path, prompting on first run."""
    config_manager = ConfigManager(CONFIG_FILE)
    path = config_manager.get_outlook_json_path()
    
    if not path:
        print("\n--- First Run Configuration ---")
        raw_path = input(r"Path (e.g., C:\Users\Name\OneDrive\OutlookSnapshot.json): ")
        clean_path = raw_path.strip(' "\'')
        config_manager.set_outlook_json_path(clean_path)
        logging.info(f"Configuration saved to {CONFIG_FILE}.")
        return clean_path
    
    return path


def get_google_service(timezone: str = 'Pacific/Auckland'):
    """Initialize and return authenticated Google Calendar service."""
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
            
    return build('calendar', 'v3', credentials=creds)


def load_outlook_snapshot(file_path: str) -> List[Dict]:
    """Load and parse the Outlook snapshot JSON file."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"Outlook snapshot file not found: {file_path}")
    
    with open(file_path, 'r', encoding='utf-8-sig') as f:
        events = json.load(f)
    
    return events if isinstance(events, list) else []


def perform_sync(service, outlook_events: List[Dict], timezone: str = 'Pacific/Auckland') -> Dict[str, int]:
    """
    Perform the synchronization between Outlook and Google Calendar.
    Returns a dictionary with sync statistics: {'created': int, 'updated': int, 'deleted': int}
    """
    # Index Outlook events by ExchangeID
    outlook_dict = {ev.get('ExchangeID'): ev for ev in outlook_events if ev.get('ExchangeID')}
    
    # Fetch Google Calendar events
    logging.info("Fetching Google Calendar events...")
    now = datetime.datetime.now(datetime.timezone.utc).isoformat().replace('+00:00', 'Z')
    events_result = service.events().list(calendarId='primary', timeMin=now,
                                          singleEvents=True, maxResults=2500).execute()
    google_events = events_result.get('items', [])
    
    # Index Google events by ExchangeID
    google_dict = {}
    for ge in google_events:
        props = ge.get('extendedProperties', {}).get('private', {})
        exchange_id = props.get('ExchangeID')
        if exchange_id:
            google_dict[exchange_id] = ge
    
    logging.info(f"Found {len(outlook_dict)} events in Outlook, {len(google_dict)} tracked events in Google.")
    logging.info("Syncing events...")
    
    stats = {'created': 0, 'updated': 0, 'deleted': 0}
    
    # UPSERT: Create or update Outlook events in Google
    for ex_id, ex_ev in outlook_dict.items():
        start_time = ex_ev['StartTime'][:19]
        end_time = ex_ev['EndTime'][:19]
        subject = ex_ev.get('Subject', 'Busy')
        
        categories = ex_ev.get('Categories')
        description = ""
        if categories:
            cat_str = ", ".join(categories) if isinstance(categories, list) else str(categories)
            description = f"Categories: {cat_str}"
        
        body = {
            'summary': subject,
            'location': ex_ev.get('Location', ''),
            'description': description,
            'extendedProperties': {'private': {'ExchangeID': ex_id}}
        }
        
        if ex_ev.get('IsAllDay'):
            body['start'] = {'date': start_time[:10]}
            body['end'] = {'date': end_time[:10]}
        else:
            body['start'] = {'dateTime': start_time, 'timeZone': timezone}
            body['end'] = {'dateTime': end_time, 'timeZone': timezone}
        
        if ex_id in google_dict:
            g_id = google_dict[ex_id]['id']
            service.events().update(calendarId='primary', eventId=g_id, body=body).execute()
            del google_dict[ex_id]
            stats['updated'] += 1
            logging.info(f"Updated: {subject}")
        else:
            service.events().insert(calendarId='primary', body=body).execute()
            stats['created'] += 1
            logging.info(f"Created: {subject}")
    
    # DELETE: Remove orphaned Google events
    for remaining_ex_id, g_ev in google_dict.items():
        summary = g_ev.get('summary', 'Unknown Event')
        service.events().delete(calendarId='primary', eventId=g_ev['id']).execute()
        stats['deleted'] += 1
        logging.info(f"Deleted: {summary}")
    
    return stats

def main():
    """Main entry point for CLI usage. Performs a single sync."""
    try:
        config_manager = ConfigManager(CONFIG_FILE)
        outlook_json_path = get_config()
        timezone = config_manager.get_timezone()
        
        if not os.path.exists(outlook_json_path):
            logging.error(f"Cannot find the snapshot file at '{outlook_json_path}'.")
            sys.exit(1)
        
        logging.info("Loading Outlook snapshot...")
        outlook_events = load_outlook_snapshot(outlook_json_path)
        
        logging.info("Initializing Google Calendar service...")
        service = get_google_service(timezone)
        
        # Perform sync
        stats = perform_sync(service, outlook_events, timezone)
        
        # Update config with sync results
        config_manager.update_sync_result('Success', stats['created'], stats['updated'], stats['deleted'])
        
        logging.info(f"Sync complete! Created: {stats['created']}, Updated: {stats['updated']}, Deleted: {stats['deleted']}.")
        sys.exit(0)
        
    except Exception as e:
        config_manager = ConfigManager(CONFIG_FILE)
        config_manager.update_sync_result(f'Error: {str(e)}')
        logging.error(f"Sync failed: {e}", exc_info=True)
        sys.exit(1)


if __name__ == '__main__':
    main()