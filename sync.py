import datetime
import json
import os
import sys
import logging
import re
from pathlib import Path
from zoneinfo import ZoneInfo
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
    datefmt='%Y-%m-%d %H:%M:%S',
    handlers=[
        logging.FileHandler("sync.log", encoding='utf-8'),
        logging.StreamHandler(sys.stdout)
    ]
)

# Suppress a known non-actionable discovery cache compatibility message.
logging.getLogger("googleapiclient.discovery_cache").setLevel(logging.WARNING)

SCOPES = ['https://www.googleapis.com/auth/calendar']
CONFIG_FILE = 'config.json'


def _trim_fractional_seconds(iso_value: str) -> str:
    """Trim fractional seconds to microseconds (max 6 digits) for robust parsing."""
    match = re.match(r"^(.*?\.)(\d+)(Z|[+-]\d\d:\d\d)?$", iso_value)
    if not match:
        return iso_value

    prefix, frac, suffix = match.groups()
    trimmed_frac = frac[:6]
    return f"{prefix}{trimmed_frac}{suffix or ''}"


def normalize_event_datetime(raw_value: str, timezone_name: str) -> str:
    """
    Normalize input datetime text into local wall time for the configured timezone.
    Returns RFC3339-like local datetime without offset (Google receives timezone separately).
    """
    value = str(raw_value or "").strip()
    if not value:
        return value

    normalized_text = _trim_fractional_seconds(value)
    if normalized_text.endswith('Z'):
        normalized_text = normalized_text[:-1] + '+00:00'

    target_tz = resolve_timezone(timezone_name)
    dt = datetime.datetime.fromisoformat(normalized_text)

    if dt.tzinfo is None:
        # Power Automate often emits UTC clock time without an offset suffix.
        # Treat naive values as UTC, then convert into configured local timezone.
        dt = dt.replace(tzinfo=datetime.timezone.utc)
        dt = dt.astimezone(target_tz)
    else:
        dt = dt.astimezone(target_tz)

    return dt.strftime('%Y-%m-%dT%H:%M:%S')

def load_event_file(file_path: str) -> dict:
    with open(file_path, "r", encoding="utf-8") as f:
        return json.load(f)


def build_event_description(event: Dict) -> str:
    """Build a Google event description from common Outlook/Power Automate fields."""
    parts: List[str] = []

    categories = event.get("categories") or event.get("Categories")
    if categories:
        if isinstance(categories, list):
            categories_text = ", ".join(str(item).strip() for item in categories if str(item).strip())
        else:
            categories_text = str(categories).strip()
        if categories_text:
            parts.append(f"Categories: {categories_text}")

    content_value = (
        event.get("description")
        or event.get("details")
        or event.get("bodyPreview")
        or event.get("body")
        or event.get("content")
        or event.get("Description")
        or event.get("Details")
        or event.get("BodyPreview")
        or event.get("Body")
        or event.get("Content")
    )

    if isinstance(content_value, dict):
        content_value = (
            content_value.get("content")
            or content_value.get("Content")
            or content_value.get("body")
            or content_value.get("Body")
        )

    if content_value:
        content_text = str(content_value).strip()
        if content_text:
            parts.append(content_text)

    return "\n\n".join(parts)

def sync_single_event(service, event: dict,
                      timezone='Pacific/Auckland',
                      calendar_id='primary'):

    outlook_id = event["eventId"]
    action = event["actionType"]
    subject = str(event.get("subject", "")).strip()
    start_raw = str(event.get("start", "")).strip()
    end_raw = str(event.get("end", "")).strip()
    start_synced = start_raw
    end_synced = end_raw

    # Find existing Google event by current key first, then legacy key.
    result = service.events().list(
        calendarId=calendar_id,
        privateExtendedProperty=f"ExchangeID={outlook_id}",
        maxResults=1
    ).execute()

    matches = result.get("items", [])
    google_event = matches[0] if matches else None

    if not google_event:
        legacy_result = service.events().list(
            calendarId=calendar_id,
            privateExtendedProperty=f"OutlookID={outlook_id}",
            maxResults=1
        ).execute()
        legacy_matches = legacy_result.get("items", [])
        google_event = legacy_matches[0] if legacy_matches else None

    # DELETE
    if action == "deleted":

        google_event_id = None
        if google_event:
            google_event_id = google_event.get("id")
            service.events().delete(
                calendarId=calendar_id,
                eventId=google_event["id"]
            ).execute()

        return {
            "google_action": "deleted",
            "google_result": "deleted" if google_event else "not_found",
            "outlook_action": action,
            "subject": subject,
            "start_raw": start_raw,
            "end_raw": end_raw,
            "start_synced": start_synced,
            "end_synced": end_synced,
            "outlook_event_id": outlook_id,
            "google_event_id": google_event_id,
        }

    # CREATE / UPDATE

    body = {
        "summary": event.get("subject", ""),
        "location": event.get("location", ""),
        "extendedProperties": {
            "private": {
                "ExchangeID": outlook_id,
                "OutlookID": outlook_id,
                "Source": "Outlook"
            }
        }
    }

    description = build_event_description(event)
    if description:
        body["description"] = description

    if str(event.get("isAllDay")).lower() == "true":

        body["start"] = {
            "date": event["start"][:10]
        }

        body["end"] = {
            "date": event["end"][:10]
        }

    else:
        start_synced = normalize_event_datetime(event.get("start", ""), timezone)
        end_synced = normalize_event_datetime(event.get("end", ""), timezone)

        body["start"] = {
            "dateTime": start_synced,
            "timeZone": timezone
        }

        body["end"] = {
            "dateTime": end_synced,
            "timeZone": timezone
        }

    if google_event:

        service.events().update(
            calendarId=calendar_id,
            eventId=google_event["id"],
            body=body
        ).execute()

        return {
            "google_action": "updated",
            "outlook_action": action,
            "subject": subject,
            "start_raw": start_raw,
            "end_raw": end_raw,
            "start_synced": start_synced,
            "end_synced": end_synced,
            "outlook_event_id": outlook_id,
            "google_event_id": google_event.get("id"),
        }

    else:

        created_event = service.events().insert(
            calendarId=calendar_id,
            body=body
        ).execute()

        return {
            "google_action": "created",
            "outlook_action": action,
            "subject": subject,
            "start_raw": start_raw,
            "end_raw": end_raw,
            "start_synced": start_synced,
            "end_synced": end_synced,
            "outlook_event_id": outlook_id,
            "google_event_id": created_event.get("id"),
        }


def process_event_folder(
    service,
    folder_path: str,
    timezone: str = 'Pacific/Auckland',
    calendar_id: str = 'primary',
    delete_on_success: bool = True,
    recursive: bool = True,
    max_passes: int = 5,
) -> Dict[str, int]:
    """Process one JSON file per event from a folder and optionally delete successful files."""
    stats = {
        'created': 0,
        'updated': 0,
        'deleted': 0,
        'failed': 0,
        'processed': 0,
    }

    folder = Path(folder_path)

    def get_pending_files() -> List[Path]:
        iterator = folder.rglob('*.json') if recursive else folder.glob('*.json')
        return sorted(path for path in iterator if path.is_file())

    for pass_index in range(1, max_passes + 1):
        pending_files = get_pending_files()
        if not pending_files:
            if pass_index == 1:
                logging.info(f"No JSON event files found in {folder_path}")
            break

        logging.info(f"Processing pass {pass_index} with {len(pending_files)} JSON file(s) in {folder_path}")

        processed_this_pass = 0
        for json_file in pending_files:
            try:
                event = load_event_file(str(json_file))
                result = sync_single_event(service, event, timezone, calendar_id)
                google_action = result.get('google_action', 'unknown')

                if google_action in stats:
                    stats[google_action] += 1
                stats['processed'] += 1
                processed_this_pass += 1

                logging.info(
                    "Processed %s | outlook_action=%s | google_action=%s | google_result=%s | subject=%r | "
                    "start_raw=%s | end_raw=%s | start_synced=%s | end_synced=%s | "
                    "outlook_event_id=%s | google_event_id=%s",
                    json_file.name,
                    result.get('outlook_action', ''),
                    google_action,
                    result.get('google_result', ''),
                    result.get('subject', ''),
                    result.get('start_raw', ''),
                    result.get('end_raw', ''),
                    result.get('start_synced', ''),
                    result.get('end_synced', ''),
                    result.get('outlook_event_id', ''),
                    result.get('google_event_id', ''),
                )

                if delete_on_success:
                    json_file.unlink()
                    logging.info(f"Deleted {json_file.name} after successful sync")

            except Exception as ex:
                stats['failed'] += 1
                logging.error(f"Failed processing {json_file.name}: {ex}")

        if processed_this_pass == 0:
            break

        remaining_files = get_pending_files()
        if not remaining_files:
            break

        logging.info(f"{len(remaining_files)} JSON file(s) remain after pass {pass_index}; running another pass")

    return stats

def resolve_timezone(timezone_name: str):
    """Resolve timezone safely across environments (Windows may lack tzdata)."""
    try:
        return ZoneInfo(timezone_name)
    except Exception:
        try:
            import pytz
            return pytz.timezone(timezone_name)
        except Exception:
            logging.warning(
                f"Timezone '{timezone_name}' could not be resolved; falling back to local system timezone."
            )
            return datetime.datetime.now().astimezone().tzinfo or datetime.timezone.utc


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


def list_google_calendars(service) -> List[Dict[str, str]]:
    """Return available calendars for the authenticated user."""
    calendars = []
    page_token = None

    while True:
        result = service.calendarList().list(pageToken=page_token).execute()
        for item in result.get('items', []):
            calendars.append({
                'id': item.get('id', ''),
                'summary': item.get('summary', item.get('id', 'Unknown Calendar')),
                'primary': bool(item.get('primary', False)),
            })

        page_token = result.get('nextPageToken')
        if not page_token:
            break

    return calendars


def list_tagged_google_events(
    service,
    calendar_id: str = 'primary',
    time_min: Optional[str] = None,
    time_max: Optional[str] = None,
) -> List[Dict]:
    """
    Return Google events tagged as synced by this tool.
    Tagged events contain at least one of:
    - extendedProperties.private.ExchangeID
    - extendedProperties.private.OutlookID
    - extendedProperties.private.Source == 'Outlook'
    """
    events = []
    page_token = None

    while True:
        query = {
            'calendarId': calendar_id,
            'singleEvents': True,
            'showDeleted': False,
            'maxResults': 2500,
            'pageToken': page_token,
        }
        if time_min:
            query['timeMin'] = time_min
        if time_max:
            query['timeMax'] = time_max

        result = service.events().list(**query).execute()
        for item in result.get('items', []):
            props = item.get('extendedProperties', {}).get('private', {})
            if props.get('ExchangeID') or props.get('OutlookID') or props.get('Source') == 'Outlook':
                events.append(item)

        page_token = result.get('nextPageToken')
        if not page_token:
            break

    return events


def purge_synced_events(
    service,
    calendar_id: str = 'primary',
    time_min: Optional[str] = None,
    time_max: Optional[str] = None,
    dry_run: bool = False,
) -> Dict[str, int]:
    """Delete Google events that were tagged as synced by this tool."""
    tagged_events = list_tagged_google_events(
        service,
        calendar_id=calendar_id,
        time_min=time_min,
        time_max=time_max,
    )

    deleted = 0
    for event in tagged_events:
        event_id = event.get('id')
        if not event_id:
            continue

        summary = event.get('summary', '')
        start_obj = event.get('start', {})
        start_value = start_obj.get('dateTime') or start_obj.get('date') or ''

        if dry_run:
            logging.info(
                "DRY RUN purge candidate | event_id=%s | summary=%r | start=%s",
                event_id,
                summary,
                start_value,
            )
            continue

        service.events().delete(calendarId=calendar_id, eventId=event_id).execute()
        deleted += 1
        logging.info(
            "Purged synced event | event_id=%s | summary=%r | start=%s",
            event_id,
            summary,
            start_value,
        )

    return {
        'matched': len(tagged_events),
        'deleted': deleted,
        'dry_run': 1 if dry_run else 0,
    }


def perform_sync(service, outlook_events: List[Dict], timezone: str = 'Pacific/Auckland', calendar_id: str = 'primary') -> Dict[str, int]:
    """
    Perform the synchronization between Outlook and Google Calendar.
    Returns a dictionary with sync statistics: {'created': int, 'updated': int, 'deleted': int}
    """
    # Only sync future events (start time after now), but keep all ExchangeIDs for safe deletes.
    tz = resolve_timezone(timezone)
    now_dt = datetime.datetime.now(tz)
    future_events = []
    all_exchange_ids = set()
    for ev in outlook_events:
        exid = ev.get('ExchangeID')
        if exid:
            all_exchange_ids.add(exid)
        try:
            raw_start = str(ev['StartTime']).strip()
            # Handle ISO8601 snapshots with or without timezone (including trailing Z).
            start = datetime.datetime.fromisoformat(raw_start.replace('Z', '+00:00'))
            if start.tzinfo is None:
                start = start.replace(tzinfo=tz)
            else:
                start = start.astimezone(tz)
            if start > now_dt:
                future_events.append(ev)
        except Exception:
            continue
    # Index only future events by ExchangeID
    outlook_dict = {ev.get('ExchangeID'): ev for ev in future_events if ev.get('ExchangeID')}
    logging.info(f"Filtered {len(outlook_dict)} future events from {len(all_exchange_ids)} snapshot events.")
    
    # Fetch Google Calendar events
    logging.info("Fetching Google Calendar events...")
    now = datetime.datetime.now(datetime.timezone.utc).isoformat().replace('+00:00', 'Z')
    events_result = service.events().list(calendarId=calendar_id, timeMin=now,
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
        description_parts = []
        if categories:
            cat_str = ", ".join(categories) if isinstance(categories, list) else str(categories)
            description_parts.append(f"Categories: {cat_str}")

        # Optional richer content from snapshot (if included by Power Automate mapping).
        content_value = ex_ev.get('Content') or ex_ev.get('BodyPreview') or ex_ev.get('Body')
        if isinstance(content_value, dict):
            content_value = content_value.get('content') or content_value.get('Content')
        if content_value:
            clean_content = str(content_value).strip()
            if clean_content:
                description_parts.append(clean_content)

        description = "\n\n".join(description_parts)
        
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
            # Compare fields to avoid unnecessary updates
            def event_fields(ev):
                return (
                    ev.get('summary'),
                    ev.get('location'),
                    ev.get('description'),
                    ev.get('start'),
                    ev.get('end'),
                )
            google_ev = google_dict[ex_id]
            changed = event_fields(google_ev) != (
                body.get('summary'),
                body.get('location'),
                body.get('description'),
                body.get('start'),
                body.get('end'),
            )
            if changed:
                service.events().update(calendarId=calendar_id, eventId=g_id, body=body).execute()
                stats['updated'] += 1
                logging.info(f"Updated: {subject}")
            else:
                logging.info(f"No change: {subject}")
            del google_dict[ex_id]
        else:
            # Fallback: check historical items by ExchangeID to prevent duplicates
            # when the primary fetch window (timeMin=now) excludes past events.
            existing_result = service.events().list(
                calendarId=calendar_id,
                privateExtendedProperty=f"ExchangeID={ex_id}",
                singleEvents=True,
                maxResults=1,
            ).execute()
            existing_items = existing_result.get('items', [])
            existing_event = next((item for item in existing_items if item.get('status') != 'cancelled'), None)

            if existing_event:
                def event_fields(ev):
                    return (
                        ev.get('summary'),
                        ev.get('location'),
                        ev.get('description'),
                        ev.get('start'),
                        ev.get('end'),
                    )
                changed = event_fields(existing_event) != (
                    body.get('summary'),
                    body.get('location'),
                    body.get('description'),
                    body.get('start'),
                    body.get('end'),
                )
                if changed:
                    service.events().update(calendarId=calendar_id, eventId=existing_event['id'], body=body).execute()
                    stats['updated'] += 1
                    logging.info(f"Updated (historical match): {subject}")
                else:
                    logging.info(f"No change (historical match): {subject}")
            else:
                service.events().insert(calendarId=calendar_id, body=body).execute()
                stats['created'] += 1
                logging.info(f"Created: {subject}")
    
    # DELETE: Remove orphaned Google events, but only if ExchangeID is in the full Outlook snapshot
    for remaining_ex_id, g_ev in google_dict.items():
        if remaining_ex_id in all_exchange_ids:
            summary = g_ev.get('summary', 'Unknown Event')
            service.events().delete(calendarId=calendar_id, eventId=g_ev['id']).execute()
            stats['deleted'] += 1
            logging.info(f"Deleted: {summary}")
    
    return stats

def main():
    """Main entry point for CLI usage. Performs a single sync."""
    try:
        config_manager = ConfigManager(CONFIG_FILE)
        outlook_json_path = get_config()
        timezone = config_manager.get_timezone()
        calendar_id = config_manager.get_google_calendar_id()
        
        if not os.path.exists(outlook_json_path):
            logging.error(f"Cannot find the snapshot file at '{outlook_json_path}'.")
            sys.exit(1)
        
        logging.info("Loading Outlook snapshot...")
        outlook_events = load_outlook_snapshot(outlook_json_path)
        
        logging.info("Initializing Google Calendar service...")
        service = get_google_service(timezone)
        
        # Perform sync
        stats = perform_sync(service, outlook_events, timezone, calendar_id)
        
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