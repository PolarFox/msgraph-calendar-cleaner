import argparse
from datetime import datetime
import sys
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import aiohttp
import asyncio

import pytz

SCOPES = ['https://www.googleapis.com/auth/calendar']

class GoogleCalendarCleaner:
    def __init__(self, calendar_name=None):
        self.creds = self.authenticate()
        self.calendar_id = self.get_calendar_id(calendar_name)

    def authenticate(self):
        creds = None
        # The file token.json stores the user's access and refresh tokens
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        # If there are no (valid) credentials, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
        return creds

    def get_calendar_id(self, calendar_name):
        service = self.build_service()
        calendar_list = service.calendarList().list().execute()
        for calendar in calendar_list['items']:
            if calendar['summary'] == calendar_name:
                return calendar['id']
        print(f"Calendar '{calendar_name}' not found. Using primary calendar.")
        return 'primary'

    def fetch_events(self, start_time_iso, end_time_iso):
        service = self.build_service()
        events_result = service.events().list(calendarId=self.calendar_id, timeMin=start_time_iso,
                                              timeMax=end_time_iso, maxResults=2500, singleEvents=True,
                                              orderBy='startTime').execute()
        events = events_result.get('items', [])
        print(f"Found {len(events)} events to delete.")
        return events

    def build_service(self):
        from googleapiclient.discovery import build
        return build('calendar', 'v3', credentials=self.creds)

    def get_headers(self):
        return {
            'Authorization': f'Bearer {self.creds.token}',
            'Accept': 'application/json'
        }

    async def delete_event(self, session, event_id, semaphore):
        delete_url = f'https://www.googleapis.com/calendar/v3/calendars/{self.calendar_id}/events/{event_id}'
        async with semaphore:
            async with session.delete(delete_url, headers=self.get_headers(), ssl=True) as response:
                if response.status == 204:
                    pass
                else:
                    print(f"Could not delete event {event_id}: {response.status} {await response.text()}")
                await asyncio.sleep(0.01)

    async def delete_events(self, events):
        semaphore = asyncio.Semaphore(3)
        async with aiohttp.ClientSession() as session:
            tasks = [self.delete_event(session, event['id'], semaphore) for event in events]
            await asyncio.gather(*tasks)

    def get_headers(self):
        return {
            'Authorization': f'Bearer {self.creds.token}',
            'Content-Type': 'application/json'
        }
    
    @staticmethod
    def clean_token_cache():
        os.remove('token.json')
        print("Token cache cleaned.")


def parse_args():
    parser = argparse.ArgumentParser(description='Delete Google Calendar events within a specified time range.')
    parser.add_argument('--start', help='Start time in format YYYY-MM-DD HH:MM')
    parser.add_argument('--end', help='End time in format YYYY-MM-DD HH:MM')
    parser.add_argument('--calendar', required=False, help='Calendar name, default is primary calendar', default='primary')
    parser.add_argument('--clean', action='store_true', help='Clean token cache')
    parser.add_argument('--timezone', required=False, default='Europe/Helsinki', help='Timezone, default is Europe/Helsinki')
    return parser.parse_args()


def main():
    args = parse_args()

    if args.clean:
        GoogleCalendarCleaner.clean_token_cache()
        sys.exit

    if not args.start or not args.end:
        print('--start and --end are required')
        sys.exit(1)

    try:
        timezone = pytz.timezone(args.timezone)
    except pytz.UnknownTimeZoneError:
        print(f"Unknown timezone: {args.timezone}")
        sys.exit(1)

    try:
        start_time_naive = datetime.strptime(args.start, '%Y-%m-%d %H:%M')
        end_time_naive = datetime.strptime(args.end, '%Y-%m-%d %H:%M')
        start_time = timezone.localize(start_time_naive).astimezone(pytz.utc)
        end_time = timezone.localize(end_time_naive).astimezone(pytz.utc)
    except ValueError as e:
        print(f"Error parsing times: {e}")
        sys.exit(1)

    start_time_iso = start_time.isoformat().replace("+00:00", "Z")
    end_time_iso = end_time.isoformat().replace("+00:00", "Z")

    cleaner = GoogleCalendarCleaner(calendar_name=args.calendar)
    events = cleaner.fetch_events(start_time_iso, end_time_iso)
    asyncio.run(cleaner.delete_events(events))


if __name__ == '__main__':
    main()
