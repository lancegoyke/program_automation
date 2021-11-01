"""
A script that copies a program from a template into multiple
client spreadsheets.

Usage: change constants according to what is desired, then run

Future enhancements:
* Hide the old program (store last added sheet, lookup, then hide)
* CLI with multiple actions:
    * Copy new program to all active clients
    * Copy base sheet for new client
"""

from __future__ import print_function
from collections import namedtuple
from os import name
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

# If modifying these scopes, delete the file token.json
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]

# You'll probably want to update this
PROGRAM_NAME = "201"

# These can probably stay the same
DATA_SPREADSHEET_ID = "1tu0jNOpXEqCeEN4UKvk_Av5DE46CPNCjXBjDYZ6jhHQ"
DATA_TEMPLATE_SHEET_ID = 1880812861
DATA_CLIENT_SHEET_ID = 0
DATA_PROGRAMS_RANGE = "Programs!A2:C"
DATA_CLIENTS_RANGE = "Client Spreadsheets!A2:B"

# This is for the `test_print()` function
SPREADSHEET_ID = "1tu0jNOpXEqCeEN4UKvk_Av5DE46CPNCjXBjDYZ6jhHQ"
RANGE_NAME = "Client Spreadsheets!A2:B"


def copy(source, destination):
    """Copy one sheet to another
    source_spreadsheet (required) - spreadsheet ID
    source_sheet (required) - sheet name
    Destination (required) - spreadsheet ID
    """

    service = build("sheets", "v4", credentials=get_creds())

    result = (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=DATA_SPREADSHEET_ID, range=DATA_PROGRAMS_RANGE)
        .execute()
    )
    values = result.get("values", [])

    for row in values:
        if row[0] == source:
            template_info = row

    source_spreadsheet = row[1]
    source_sheet = row[2]

    copy_sheet_to_another_spreadsheet_request_body = {
        "destination_spreadsheet_id": destination,
    }

    # Copy the sheet
    request = (
        service.spreadsheets()
        .sheets()
        .copyTo(
            spreadsheetId=source_spreadsheet,
            sheetId=source_sheet,
            body=copy_sheet_to_another_spreadsheet_request_body,
        )
    )
    response = request.execute()
    destination_sheet = response.get("sheetId")

    # Rename the sheet
    batch_update_spreadsheet_request_body = {
        # A list of updates to apply to the spreadsheet.
        # Requests will be applied in the order they are specified.
        # If any request is not valid, no requests will be applied.
        "requests": [
            {
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": destination_sheet,
                        "title": source,
                    },
                    "fields": "title",
                }
            }
        ],
    }

    request = service.spreadsheets().batchUpdate(
        spreadsheetId=destination, body=batch_update_spreadsheet_request_body
    )
    response = request.execute()

    print(f'SUCCESS: copied sheet "{source}"')


def get_clients():
    """Return as list of client names and their spreadsheet ID"""
    Client = namedtuple("Client", ["client_name", "spreadsheet_id"])

    service = build("sheets", "v4", credentials=get_creds())

    # get the IDs from my Data.Client Spreadsheets sheet
    result = (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=DATA_SPREADSHEET_ID, range=DATA_CLIENTS_RANGE)
        .execute()
    )
    values = result.get("values", [])

    return [Client(row[0], row[1]) for row in values]


def get_creds():
    """Get connected to Google account"""
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return creds


def print_test():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    service = build("sheets", "v4", credentials=get_creds())

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = (
        sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
    )
    values = result.get("values", [])

    if not values:
        print("No data found.")
    else:
        client_sheets = [row[1] for row in values]
        print(client_sheets)


def main():

    for client in get_clients():
        copy(PROGRAM_NAME, client.spreadsheet_id)


if __name__ == "__main__":
    main()
