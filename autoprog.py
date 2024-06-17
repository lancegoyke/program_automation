"""
A script that copies a program from a template into multiple
client spreadsheets.

Usage: change constants according to what is desired, then run

TODO:
- [x] get_template_programs()
- [ ] run with test flag

Future enhancements:
* Hide the old program (store last added sheet, lookup, then hide)
* CLI with multiple actions:
    * Copy new program to all active clients
    * Copy base sheet for new client
"""

from __future__ import print_function
from collections import namedtuple
from datetime import datetime
import os.path
from googleapiclient.discovery import build, Resource
from googleapiclient.errors import HttpError
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

# If modifying these scopes, delete the file token.json
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/spreadsheets",
]

# You'll probably want to update this
PROGRAM_NAME = "202"

# These can probably stay the same
DATA_SPREADSHEET_ID = "1tu0jNOpXEqCeEN4UKvk_Av5DE46CPNCjXBjDYZ6jhHQ"
DATA_TEMPLATE_SHEET_ID = 1880812861
DATA_CLIENT_SHEET_ID = 0
DATA_PROGRAMS_RANGE = "Programs!A2:C"
DATA_CLIENTS_RANGE = "Client Spreadsheets!A2:B"

# This is for the `test_print()` function
SPREADSHEET_ID = "1tu0jNOpXEqCeEN4UKvk_Av5DE46CPNCjXBjDYZ6jhHQ"
RANGE_NAME = "Client Spreadsheets!A2:B"


def get_template_programs(service: Resource) -> list[list[str]]:
    """
    Fetches a spreadsheet full of template programs worth copying.

    Args:
        service (Resource): The Google API service object.

    Returns:
        List[List[str]]: A list of lists containing the template programs.

    Example:
        [
          # Program Name  Spreadsheet ID                                 Sheet ID
            ["201",      "testjNOpXEqCeEN4UKvk_Av5DE46CPNCjXBjDYZ6jhHQ", "1111812861"],
            ["202",      "testjNOpXEqCeEN4UKvk_Av5DE46CPNCjXBjDYZ6jhHQ", "1111812861"],
            ["203",      "testjNOpXEqCeEN4UKvk_Av5DE46CPNCjXBjDYZ6jhHQ", "1111812861"],
            ["204",      "testjNOpXEqCeEN4UKvk_Av5DE46CPNCjXBjDYZ6jhHQ", "1111812861"],
            # ...
        ]
    """
    return (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=DATA_SPREADSHEET_ID, range=DATA_PROGRAMS_RANGE)
        .execute()
        .get("values", [])
    )


def copy(service: Resource, program_name: str, destination: str):
    """
    Copy one sheet to a different spreadsheet
    `program_name: str` (required) - the program name
    `destination: str` (required) - spreadsheet ID
    """

    # Get the template programs
    data_programs = get_template_programs(service)

    for row in data_programs:
        if row[0] == program_name:
            template_info = row

    source_spreadsheet: str = template_info[1]
    source_sheet: int = template_info[2]

    copy_sheet_to_another_spreadsheet_request_body: dict = {
        "destination_spreadsheet_id": destination,
    }

    # Copy the sheet
    # https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.sheets/copyTo
    # https://googleapis.github.io/google-api-python-client/docs/dyn/sheets_v4.spreadsheets.sheets.html
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
    program_name_with_MM_YY = f"{program_name} - {datetime.now().strftime('%m/%y')}"
    batch_update_spreadsheet_request_body = {
        # A list of updates to apply to the spreadsheet.
        # Requests will be applied in the order they are specified.
        # If any request is not valid, no requests will be applied.
        "requests": [
            {
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": destination_sheet,
                        "title": program_name_with_MM_YY,
                    },
                    "fields": "title",
                }
            }
        ],
        "includeSpreadsheetInResponse": True,
    }
    request = service.spreadsheets().batchUpdate(
        spreadsheetId=destination, body=batch_update_spreadsheet_request_body
    )
    response: dict = request.execute()

    try:
        print(
            f'SUCCESS: copied "{program_name_with_MM_YY}" sheet to {response["updatedSpreadsheet"]["properties"]["title"]}'
        )
    except Exception as e:
        print("There was an error printing the sheet properties")
        print(e)
        print(f'SUCCESS: copied sheet "{program_name_with_MM_YY}')


def get_clients(service: Resource):
    """
    Return a list of client names and their spreadsheet ID
    """
    Client = namedtuple("Client", ["client_name", "spreadsheet_id"])

    # service = build("sheets", "v4", credentials=get_creds())

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
        print("Found 'token.json' file")
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
    with build("sheets", "v4", credentials=get_creds()) as service:
        for client in get_clients(service):
            copy(service, PROGRAM_NAME, client.spreadsheet_id)


if __name__ == "__main__":
    main()
