import asyncio
from time import sleep
import requests
import argparse
import os
import sys
import dotenv as dotenv
import msal
import pytz
import json
from datetime import datetime
import aiohttp
import ssl

dotenv.load_dotenv()

class CalendarCleaner:
    def __init__(self, client_id, tenant_id, timezone='Europe/Helsinki'):
        self.client_id = client_id
        self.tenant_id = tenant_id
        self.timezone = timezone
        self.cache = self.load_cache()
        self.app = msal.PublicClientApplication(
            client_id, 
            authority=f"https://login.microsoftonline.com/{tenant_id}", 
            token_cache=self.cache
        )
        self.scopes = ["Calendars.ReadWrite"]
        self.token = self.acquire_token()
        self.headers = {
            'Authorization': f'Bearer {self.token}',
            'Content-Type': 'application/json'
        }

    def load_cache(self):
        cache = msal.SerializableTokenCache()
        if os.path.exists("token_cache.bin"):
            cache.deserialize(open("token_cache.bin", "r").read())
        return cache

    def save_cache(self):
        if self.cache.has_state_changed:
            with open("token_cache.bin", "w") as f:
                f.write(self.cache.serialize())

    def clean_cache(self):
        if os.path.exists("token_cache.bin"):
            os.remove("token_cache.bin")

    def acquire_token(self):
        accounts = self.app.get_accounts()
        if accounts:
            result = self.app.acquire_token_silent(self.scopes, account=accounts[0])
        else:
            flow = self.app.initiate_device_flow(scopes=self.scopes)
            if "user_code" not in flow:
                raise Exception("Failed to obtain device code.")
            print(flow["message"])
            sys.stdout.flush()
            result = self.app.acquire_token_by_device_flow(flow)

        if "access_token" in result:
            self.save_cache()
            return result['access_token']
        else:
            raise Exception(f"Could not acquire access token: {result.get('error')}")

    def fetch_events(self, start_time_iso, end_time_iso):
        events = []
        next_link = f'https://graph.microsoft.com/v1.0/me/calendarview?startDateTime={start_time_iso}&endDateTime={end_time_iso}&$top=500'
        print("Fetching calendar events...")

        while next_link:
            if len(events) > 9500:
                print(f"Fetched {len(events)} events, stopping.")
                break
            try:
                print(f"Fetching link: {next_link}")
                response = requests.get(next_link, headers=self.headers, verify=True)
                response.raise_for_status()
                data = response.json()
                events.extend(data['value'])
                next_link = data.get('@odata.nextLink')
                sleep(0.05)
            except requests.exceptions.HTTPError as err:
                print(f"Error fetching calendar events: {err}")
                print(response.text)
                sys.exit(1)

        print(f"Found {len(events)} events to delete.")
        return events

    async def delete_event(self, session, event_id, semaphore):
        delete_url = f'https://graph.microsoft.com/v1.0/me/events/{event_id}'
        async with semaphore:
            async with session.delete(delete_url, headers=self.headers, ssl=True) as response:
                if response.status == 204:
                    print(f"Deleted event {event_id}")
                else:
                    print(f"Could not delete event {event_id}: {response.status} {await response.text()}")
                await asyncio.sleep(0.05)

    async def delete_events(self, events):
        semaphore = asyncio.Semaphore(3)
        async with aiohttp.ClientSession() as session:
            tasks = [self.delete_event(session, event['id'], semaphore) for event in events]
            await asyncio.gather(*tasks)

def parse_args():
    parser = argparse.ArgumentParser(description='Delete calendar events within a specified time range.')
    parser.add_argument('--start', help='Start time in format YYYY-MM-DD HH:MM')
    parser.add_argument('--end', help='End time in format YYYY-MM-DD HH:MM')
    parser.add_argument('--timezone', required=False, default='Europe/Helsinki', help='Timezone, default is Europe/Helsinki')
    parser.add_argument('--clean', action='store_true', help='Clean token cache')
    return parser.parse_args()

def main():
    args = parse_args()

    if not args.clean and (not args.start or not args.end):
        print('--start and --end are required unless --clean is specified')
        sys.exit(1)

    if args.clean:
        CalendarCleaner.clean_cache()
        print("Token cache cleaned.")
        sys.exit(0)

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

    client_id = os.environ.get('CLIENT_ID') or input('Enter client ID: ')
    tenant_id = os.environ.get('TENANT_ID') or input('Enter tenant ID: ')

    cleaner = CalendarCleaner(client_id, tenant_id, args.timezone)
    events = cleaner.fetch_events(start_time_iso, end_time_iso)
    asyncio.run(cleaner.delete_events(events))

if __name__ == '__main__':
    main()