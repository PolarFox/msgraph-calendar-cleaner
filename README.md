# Calendar entry removal via Microsoft Graph API

IF you ever happend to create a loop with your calendar entries with low-code tools like Power Automate, you might want to remove them all at once.
This happened to me and I created this script to remove all the entries from my calendar in specific time slots.

This script uses Microsoft Graph API to authenticate and remove the calendar entries.

# Features

- Microsoft Graph API cleaner for calendar entries
- Google Calendar API cleaner for calendar entries
- Remove calendar entries from a specific time slot
- Token caching for faster authentication on subsequent runs
- .env file support for environment variables

## Prerequisites

- Python 3.8 or later
- A Microsoft 365 account with a mailbox OR Google account with calendar and Developer API enabled
- A registered Azure AD application with the necessary permissions OR Google Cloud project with the necessary permissions
- The application ID and secret

## Other

Application uses device code flow to authenticate user.

## How to run

1. Install the required packages:

```bash
pip install -r requirements.txt
```

2. Run the script:

```bash
python msgraph_cleaner.py --start "YYYY-MM-DD HH:MM" --end "YYYY-MM-DD HH:MM"
```

OR

```bash 
python google_cleaner.py --start "YYYY-MM-DD HH:MM" --end "YYYY-MM-DD HH:MM" --calendar "calendar name"
```

Replace `YYYY-MM-DD HH:MM` with the start and end date and time of the calendar entry you want to remove.
Script will prompt the application details if they are not found from environment variables.

Environment variables:
* `CLIENT_ID` - Azure AD application ID
* `CLIENT_SECRET` - Azure AD application secret
* `TENANT_ID` - Azure AD tenant ID

3. Follow the instructions on the screen to authenticate the user and remove the calendar entry.

4. Clean up the token cache

For security reasons, CLEAN UP THE TOKEN CACHE after you have finished using the script.

To clean up the token cache, run the script with the `--clean` option

```bash
python msgraph_cleaner.py --clean
```
OR
```bash
python google_cleaner.py --clean
```