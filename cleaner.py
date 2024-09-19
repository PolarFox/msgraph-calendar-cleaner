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

dotenv.load_dotenv()


def load_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists("token_cache.bin"):
        cache.deserialize(open("token_cache.bin", "r").read())
    return cache


def save_cache(cache):
    if cache.has_state_changed:
        with open("token_cache.bin", "w") as f:
            f.write(cache.serialize())


async def delete_event(session, event_id, headers, semaphore):
    """
    Delete a calendar event using the Microsoft Graph API.
    :param session: aiohttp ClientSession object
    :param event_id: Event ID
    :param headers: Request headers
    :param semaphore: Semaphore object to limit concurrent requests
    """
    delete_url = f'https://graph.microsoft.com/v1.0/me/events/{event_id}'
    async with semaphore:
        async with session.delete(delete_url, headers=headers) as response:
            if response.status == 204:
                print(f"Deleted event {event_id}")
            else:
                print(f"Could not delete event {event_id}: {response.status} {await response.text()}")
            # Wait for 50ms to avoid throttling
            await asyncio.sleep(0.05)


async def main():
    """
    Main function to delete calendar events within a specified time range.
    """
    parser = argparse.ArgumentParser(description='Delete calendar events within a specified time range.')
    parser.add_argument('--start', required=True, help='Start time in format YYYY-MM-DD HH:MM')
    parser.add_argument('--end', required=True, help='End time in format YYYY-MM-DD HH:MM')
    parser.add_argument('--timezone', required=False, default='Europe/Helsinki', help='Timezone, default is Europe/Helsinki')

    args = parser.parse_args()

    # Set timezone
    try:
        timezone = pytz.timezone(args.timezone)
    except pytz.UnknownTimeZoneError:
        print(f"Unknown timezone: {args.timezone}")
        sys.exit(1)

    # Convert times to datetime objects
    try:
        start_time_naive = datetime.strptime(args.start, '%Y-%m-%d %H:%M')
        end_time_naive = datetime.strptime(args.end, '%Y-%m-%d %H:%M')
        start_time = timezone.localize(start_time_naive).astimezone(pytz.utc)
        end_time = timezone.localize(end_time_naive).astimezone(pytz.utc)
    except ValueError as e:
        print(f"Error parsing times: {e}")
        sys.exit(1)

    # Convert to ISO 8601 format in UTC
    start_time_iso = start_time.isoformat().replace("+00:00", "Z")
    end_time_iso = end_time.isoformat().replace("+00:00", "Z")

    # Fetch credentials from environment variables or prompt the user
    client_id = os.environ.get('CLIENT_ID') or input('Enter client ID: ')
    tenant_id = os.environ.get('TENANT_ID') or input('Enter tenant ID: ')

    # Load token cache
    cache = load_cache()

    # Create MSAL application object with cache
    app = msal.PublicClientApplication(client_id, authority=f"https://login.microsoftonline.com/{tenant_id}", token_cache=cache)

    # Define required scopes
    scopes = ["Calendars.ReadWrite"]

    # Acquire token using Device Code Flow
    result = None
    accounts = app.get_accounts()
    if accounts:
        # Use existing account
        result = app.acquire_token_silent(scopes, account=accounts[0])

    if not result:
        # Prompt user to sign in with device code
        flow = app.initiate_device_flow(scopes=scopes)
        if "user_code" not in flow:
            print("Failed to obtain device code.")
            sys.exit(1)
        print(flow["message"])
        sys.stdout.flush()
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" in result:
        token = result['access_token']
        # Save the cache
        save_cache(cache)
    else:
        print(f"Could not acquire access token: {result.get('error')}")
        sys.exit(1)

    headers = {
        'Authorization': f'Bearer {token}',
        'Content-Type': 'application/json'
    }

    # Fetch calendar events using the /me endpoint
    events = []
    next_link = f'https://graph.microsoft.com/v1.0/me/calendarview?startDateTime={start_time_iso}&endDateTime={end_time_iso}&$top=500'

    print("Fetching calendar events...")

    while next_link:
        if len(events) > 9500:
            print(f"Fetched {len(events)} events, stopping.")
            break
        try:
            print(f"Fetching link: {next_link}")
            response = requests.get(next_link, headers=headers)
            response.raise_for_status()
            data = response.json()
            events.extend(data['value'])
            next_link = data.get('@odata.nextLink')
            # Wait for 50ms to avoid throttling
            sleep(0.05)
        except requests.exceptions.HTTPError as err:
            print(f"Error fetching calendar events: {err}")
            print(response.text)
            sys.exit(1)

    print(f"Found {len(events)} events to delete.")

    if True:
        # Delete calendar events using semaphore to limit concurrent requests
        semaphore = asyncio.Semaphore(3)
        async with aiohttp.ClientSession() as session:
            tasks = [delete_event(session, event['id'], headers, semaphore) for event in events]
            await asyncio.gather(*tasks)
        for event in events:
            event_id = event['id']
            delete_url = f'https://graph.microsoft.com/v1.0/me/events/{event_id}'
            del_response = requests.delete(delete_url, headers=headers)
            if del_response.status_code == 204:
                print(f"Deleted event {event_id}")
            else:
                print(f"Could not delete event {event_id}: {del_response.status_code} {del_response.text}")
            # Wait for 50ms to avoid throttling
            sleep(0.05)

if __name__ == '__main__':
    asyncio.run(main())
