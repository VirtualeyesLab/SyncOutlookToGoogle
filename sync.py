import datetime
import json
import os
import sys
import logging
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

# 1 & 2. Configure Logging (Writes to terminal AND sync.log file)
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

def get_config():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
            config = json.load(f)
            return config.get('OUTLOOK_JSON_PATH')
    
    print("\n--- First Run Configuration ---")
    raw_path = input(r"Path (e.g., C:\Users\Name\OneDrive\OutlookSnapshot.json): ")
    clean_path = raw_path.strip(' "\'')
    
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump({'OUTLOOK_JSON_PATH': clean_path}, f, indent=4)
        
    logging.info(f"Configuration saved to {CONFIG_FILE}.")
    return clean_path

def get_google_service():
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

def main():
    outlook_json_path = get_config()
    
    if not os.path.exists(outlook_json_path):
        logging.error(f"Cannot find the snapshot file at '{outlook_json_path}'.")
        sys.exit(1)

    logging.info("Loading Outlook snapshot...")
    with open(outlook_json_path, 'r', encoding='utf-8-sig') as f:
        outlook_events = json.load(f)

    outlook_dict = {ev['ExchangeID']: ev for ev in outlook_events}

    logging.info("Fetching Google Calendar events...")
    service = get_google_service()
    
    # 3. Fixed Deprecation Warning
    now = datetime.datetime.now(datetime.timezone.utc).isoformat().replace('+00:00', 'Z')
    events_result = service.events().list(calendarId='primary', timeMin=now,
                                          singleEvents=True, maxResults=2500).execute()
    google_events = events_result.get('items', [])

    google_dict = {}
    for ge in google_events:
        props = ge.get('extendedProperties', {}).get('private', {})
        exchange_id = props.get('ExchangeID')
        if exchange_id:
            google_dict[exchange_id] = ge

    logging.info(f"Found {len(outlook_dict)} events in Outlook, {len(google_dict)} tracked events in Google.")
    logging.info("Syncing events...")
    
    stats = {'created': 0, 'updated': 0, 'deleted': 0}

    # UPSERT
    for ex_id, ex_ev in outlook_dict.items():
        # 4. Timezone fix: Removed the 'Z' appended here
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

        # 4. Timezone fix: Explicitly define the local timezone
        if ex_ev.get('IsAllDay'):
            body['start'] = {'date': start_time[:10]}
            body['end'] = {'date': end_time[:10]}
        else:
            body['start'] = {'dateTime': start_time, 'timeZone': 'Pacific/Auckland'}
            body['end'] = {'dateTime': end_time, 'timeZone': 'Pacific/Auckland'}

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

    # DELETE ORPHANS
    for remaining_ex_id, g_ev in google_dict.items():
        summary = g_ev.get('summary', 'Unknown Event')
        service.events().delete(calendarId='primary', eventId=g_ev['id']).execute()
        stats['deleted'] += 1
        logging.info(f"Deleted: {summary}")

    logging.info(f"Sync complete! Created: {stats['created']}, Updated: {stats['updated']}, Deleted: {stats['deleted']}.")

if __name__ == '__main__':
    main()