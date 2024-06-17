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
PROGRAM_NAME = "203"

# These can probably stay the same
DATA_SPREADSHEET_ID = "1tu0jNOpXEqCeEN4UKvk_Av5DE46CPNCjXBjDYZ6jhHQ"
DATA_TEMPLATE_SHEET_ID = 1880812861
DATA_CLIENT_SHEET_ID = 0
DATA_PROGRAMS_RANGE = "Programs!A2:C"
# DATA_CLIENTS_RANGE = "Client Spreadsheets!A2:B"
# uncomment the following line to use testing data
DATA_CLIENTS_RANGE = "TESTDATA Client Spreadsheets!A2:B"

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


def spreadsheets_sheets_copyto(
    service: Resource,
    source_spreadsheet: str,
    source_sheet: int,
    destination: str,
) -> dict:
    """
    Copies the sheet from one spreadsheet to another.

    Args:
        service (Resource): The Google API service object.
        program_name (str): The name of the program.
        source_spreadsheet (str): The source spreadsheet ID.
        source_sheet (int): The source sheet ID.
        destination (str): The destination spreadsheet ID.

    Returns:
        dict: The properties of the newly created sheet in the following form:

        { # Properties of a sheet.
            "dataSourceSheetProperties": { # Additional properties of a DATA_SOURCE sheet. # Output only. If present, the field contains DATA_SOURCE sheet specific properties.
                "columns": [ # The columns displayed on the sheet, corresponding to the values in RowData.
                { # A column in a data source.
                    "formula": "A String", # The formula of the calculated column.
                    "reference": { # An unique identifier that references a data source column. # The column reference.
                    "name": "A String", # The display name of the column. It should be unique within a data source.
                    },
                },
                ],
                "dataExecutionStatus": { # The data execution status. A data execution is created to sync a data source object with the latest data from a DataSource. It is usually scheduled to run at background, you can check its state to tell if an execution completes There are several scenarios where a data execution is triggered to run: * Adding a data source creates an associated data source sheet as well as a data execution to sync the data from the data source to the sheet. * Updating a data source creates a data execution to refresh the associated data source sheet similarly. * You can send refresh request to explicitly refresh one or multiple data source objects. # The data execution status.
                "errorCode": "A String", # The error code.
                "errorMessage": "A String", # The error message, which may be empty.
                "lastRefreshTime": "A String", # Gets the time the data last successfully refreshed.
                "state": "A String", # The state of the data execution.
                },
                "dataSourceId": "A String", # ID of the DataSource the sheet is connected to.
            },
            "gridProperties": { # Properties of a grid. # Additional properties of the sheet if this sheet is a grid. (If the sheet is an object sheet, containing a chart or image, then this field will be absent.) When writing it is an error to set any grid properties on non-grid sheets. If this sheet is a DATA_SOURCE sheet, this field is output only but contains the properties that reflect how a data source sheet is rendered in the UI, e.g. row_count.
                "columnCount": 42, # The number of columns in the grid.
                "columnGroupControlAfter": True or False, # True if the column grouping control toggle is shown after the group.
                "frozenColumnCount": 42, # The number of columns that are frozen in the grid.
                "frozenRowCount": 42, # The number of rows that are frozen in the grid.
                "hideGridlines": True or False, # True if the grid isn't showing gridlines in the UI.
                "rowCount": 42, # The number of rows in the grid.
                "rowGroupControlAfter": True or False, # True if the row grouping control toggle is shown after the group.
            },
            "hidden": True or False, # True if the sheet is hidden in the UI, false if it's visible.
            "index": 42, # The index of the sheet within the spreadsheet. When adding or updating sheet properties, if this field is excluded then the sheet is added or moved to the end of the sheet list. When updating sheet indices or inserting sheets, movement is considered in "before the move" indexes. For example, if there were three sheets (S1, S2, S3) in order to move S1 ahead of S2 the index would have to be set to 2. A sheet index update request is ignored if the requested index is identical to the sheets current index or if the requested new index is equal to the current sheet index + 1.
            "rightToLeft": True or False, # True if the sheet is an RTL sheet instead of an LTR sheet.
            "sheetId": 42, # The ID of the sheet. Must be non-negative. This field cannot be changed once set.
            "sheetType": "A String", # The type of sheet. Defaults to GRID. This field cannot be changed once set.
            "tabColor": { # Represents a color in the RGBA color space. This representation is designed for simplicity of conversion to and from color representations in various languages over compactness. For example, the fields of this representation can be trivially provided to the constructor of `java.awt.Color` in Java; it can also be trivially provided to UIColor's `+colorWithRed:green:blue:alpha` method in iOS; and, with just a little work, it can be easily formatted into a CSS `rgba()` string in JavaScript. This reference page doesn't have information about the absolute color space that should be used to interpret the RGB value—for example, sRGB, Adobe RGB, DCI-P3, and BT.2020. By default, applications should assume the sRGB color space. When color equality needs to be decided, implementations, unless documented otherwise, treat two colors as equal if all their red, green, blue, and alpha values each differ by at most `1e-5`. Example (Java): import com.google.type.Color; // ... public static java.awt.Color fromProto(Color protocolor) { float alpha = protocolor.hasAlpha() ? protocolor.getAlpha().getValue() : 1.0; return new java.awt.Color( protocolor.getRed(), protocolor.getGreen(), protocolor.getBlue(), alpha); } public static Color toProto(java.awt.Color color) { float red = (float) color.getRed(); float green = (float) color.getGreen(); float blue = (float) color.getBlue(); float denominator = 255.0; Color.Builder resultBuilder = Color .newBuilder() .setRed(red / denominator) .setGreen(green / denominator) .setBlue(blue / denominator); int alpha = color.getAlpha(); if (alpha != 255) { result.setAlpha( FloatValue .newBuilder() .setValue(((float) alpha) / denominator) .build()); } return resultBuilder.build(); } // ... Example (iOS / Obj-C): // ... static UIColor* fromProto(Color* protocolor) { float red = [protocolor red]; float green = [protocolor green]; float blue = [protocolor blue]; FloatValue* alpha_wrapper = [protocolor alpha]; float alpha = 1.0; if (alpha_wrapper != nil) { alpha = [alpha_wrapper value]; } return [UIColor colorWithRed:red green:green blue:blue alpha:alpha]; } static Color* toProto(UIColor* color) { CGFloat red, green, blue, alpha; if (![color getRed:&red green:&green blue:&blue alpha:&alpha]) { return nil; } Color* result = [[Color alloc] init]; [result setRed:red]; [result setGreen:green]; [result setBlue:blue]; if (alpha <= 0.9999) { [result setAlpha:floatWrapperWithValue(alpha)]; } [result autorelease]; return result; } // ... Example (JavaScript): // ... var protoToCssColor = function(rgb_color) { var redFrac = rgb_color.red || 0.0; var greenFrac = rgb_color.green || 0.0; var blueFrac = rgb_color.blue || 0.0; var red = Math.floor(redFrac * 255); var green = Math.floor(greenFrac * 255); var blue = Math.floor(blueFrac * 255); if (!('alpha' in rgb_color)) { return rgbToCssColor(red, green, blue); } var alphaFrac = rgb_color.alpha.value || 0.0; var rgbParams = [red, green, blue].join(','); return ['rgba(', rgbParams, ',', alphaFrac, ')'].join(''); }; var rgbToCssColor = function(red, green, blue) { var rgbNumber = new Number((red << 16) | (green << 8) | blue); var hexString = rgbNumber.toString(16); var missingZeros = 6 - hexString.length; var resultBuilder = ['#']; for (var i = 0; i < missingZeros; i++) { resultBuilder.push('0'); } resultBuilder.push(hexString); return resultBuilder.join(''); }; // ... # The color of the tab in the UI. Deprecated: Use tab_color_style.
                "alpha": 3.14, # The fraction of this color that should be applied to the pixel. That is, the final pixel color is defined by the equation: `pixel color = alpha * (this color) + (1.0 - alpha) * (background color)` This means that a value of 1.0 corresponds to a solid color, whereas a value of 0.0 corresponds to a completely transparent color. This uses a wrapper message rather than a simple float scalar so that it is possible to distinguish between a default value and the value being unset. If omitted, this color object is rendered as a solid color (as if the alpha value had been explicitly given a value of 1.0).
                "blue": 3.14, # The amount of blue in the color as a value in the interval [0, 1].
                "green": 3.14, # The amount of green in the color as a value in the interval [0, 1].
                "red": 3.14, # The amount of red in the color as a value in the interval [0, 1].
            },
            "tabColorStyle": { # A color value. # The color of the tab in the UI. If tab_color is also set, this field takes precedence.
                "rgbColor": { # Represents a color in the RGBA color space. This representation is designed for simplicity of conversion to and from color representations in various languages over compactness. For example, the fields of this representation can be trivially provided to the constructor of `java.awt.Color` in Java; it can also be trivially provided to UIColor's `+colorWithRed:green:blue:alpha` method in iOS; and, with just a little work, it can be easily formatted into a CSS `rgba()` string in JavaScript. This reference page doesn't have information about the absolute color space that should be used to interpret the RGB value—for example, sRGB, Adobe RGB, DCI-P3, and BT.2020. By default, applications should assume the sRGB color space. When color equality needs to be decided, implementations, unless documented otherwise, treat two colors as equal if all their red, green, blue, and alpha values each differ by at most `1e-5`. Example (Java): import com.google.type.Color; // ... public static java.awt.Color fromProto(Color protocolor) { float alpha = protocolor.hasAlpha() ? protocolor.getAlpha().getValue() : 1.0; return new java.awt.Color( protocolor.getRed(), protocolor.getGreen(), protocolor.getBlue(), alpha); } public static Color toProto(java.awt.Color color) { float red = (float) color.getRed(); float green = (float) color.getGreen(); float blue = (float) color.getBlue(); float denominator = 255.0; Color.Builder resultBuilder = Color .newBuilder() .setRed(red / denominator) .setGreen(green / denominator) .setBlue(blue / denominator); int alpha = color.getAlpha(); if (alpha != 255) { result.setAlpha( FloatValue .newBuilder() .setValue(((float) alpha) / denominator) .build()); } return resultBuilder.build(); } // ... Example (iOS / Obj-C): // ... static UIColor* fromProto(Color* protocolor) { float red = [protocolor red]; float green = [protocolor green]; float blue = [protocolor blue]; FloatValue* alpha_wrapper = [protocolor alpha]; float alpha = 1.0; if (alpha_wrapper != nil) { alpha = [alpha_wrapper value]; } return [UIColor colorWithRed:red green:green blue:blue alpha:alpha]; } static Color* toProto(UIColor* color) { CGFloat red, green, blue, alpha; if (![color getRed:&red green:&green blue:&blue alpha:&alpha]) { return nil; } Color* result = [[Color alloc] init]; [result setRed:red]; [result setGreen:green]; [result setBlue:blue]; if (alpha <= 0.9999) { [result setAlpha:floatWrapperWithValue(alpha)]; } [result autorelease]; return result; } // ... Example (JavaScript): // ... var protoToCssColor = function(rgb_color) { var redFrac = rgb_color.red || 0.0; var greenFrac = rgb_color.green || 0.0; var blueFrac = rgb_color.blue || 0.0; var red = Math.floor(redFrac * 255); var green = Math.floor(greenFrac * 255); var blue = Math.floor(blueFrac * 255); if (!('alpha' in rgb_color)) { return rgbToCssColor(red, green, blue); } var alphaFrac = rgb_color.alpha.value || 0.0; var rgbParams = [red, green, blue].join(','); return ['rgba(', rgbParams, ',', alphaFrac, ')'].join(''); }; var rgbToCssColor = function(red, green, blue) { var rgbNumber = new Number((red << 16) | (green << 8) | blue); var hexString = rgbNumber.toString(16); var missingZeros = 6 - hexString.length; var resultBuilder = ['#']; for (var i = 0; i < missingZeros; i++) { resultBuilder.push('0'); } resultBuilder.push(hexString); return resultBuilder.join(''); }; // ... # RGB color. The [`alpha`](/sheets/api/reference/rest/v4/spreadsheets/other#Color.FIELDS.alpha) value in the [`Color`](/sheets/api/reference/rest/v4/spreadsheets/other#color) object isn't generally supported.
                "alpha": 3.14, # The fraction of this color that should be applied to the pixel. That is, the final pixel color is defined by the equation: `pixel color = alpha * (this color) + (1.0 - alpha) * (background color)` This means that a value of 1.0 corresponds to a solid color, whereas a value of 0.0 corresponds to a completely transparent color. This uses a wrapper message rather than a simple float scalar so that it is possible to distinguish between a default value and the value being unset. If omitted, this color object is rendered as a solid color (as if the alpha value had been explicitly given a value of 1.0).
                "blue": 3.14, # The amount of blue in the color as a value in the interval [0, 1].
                "green": 3.14, # The amount of green in the color as a value in the interval [0, 1].
                "red": 3.14, # The amount of red in the color as a value in the interval [0, 1].
                },
                "themeColor": "A String", # Theme color.
            },
            "title": "A String", # The name of the sheet.
            }

    Links:
        - https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.sheets/copyTo
        - https://googleapis.github.io/google-api-python-client/docs/dyn/sheets_v4.spreadsheets.sheets.html
    """
    return (
        service.spreadsheets()
        .sheets()
        .copyTo(
            spreadsheetId=source_spreadsheet,
            sheetId=source_sheet,
            body={"destination_spreadsheet_id": destination},
        )
        .execute()
    )


def rename_sheet(service: Resource, spreadsheet_id: str, sheet_id: int, new_title: str):
    batch_update_spreadsheet_request_body = {
        "requests": [
            {
                "updateSheetProperties": {
                    "properties": {
                        "sheetId": sheet_id,
                        "title": new_title,
                    },
                    "fields": "title",
                }
            }
        ],
        "includeSpreadsheetInResponse": True,
    }
    return (
        service.spreadsheets()
        .batchUpdate(
            spreadsheetId=spreadsheet_id, body=batch_update_spreadsheet_request_body
        )
        .execute()
    )


def copy(service: Resource, program_name: str, destination_spreadsheet_id: str):
    """
    Copy one sheet to a different spreadsheet
    `program_name: str` (required) - the program name
    `destination: str` (required) - spreadsheet ID
    """

    # Get the template programs
    try:
        data_programs = get_template_programs(service)
    except HttpError as e:
        print(f"ERROR {e.status_code}: {e.reason}")
        print("ERROR: could not get template programs")
        return

    for row in data_programs:
        if row[0] == program_name:
            template_info = row

    source_spreadsheet: str = template_info[1]
    source_sheet: int = template_info[2]

    try:
        destination_sheet = spreadsheets_sheets_copyto(
            service, source_spreadsheet, source_sheet, destination_spreadsheet_id
        ).get("sheetId")
    except HttpError as e:
        print(f"ERROR {e.status_code}: {e.reason}")
        print(f'ERROR: could not copy "{program_name}" to {destination_spreadsheet_id}')
        return

    # Rename the sheet
    new_title = f"{program_name} - {datetime.now().strftime('%m/%y')}"
    try:
        updated_spreadsheet = rename_sheet(
            service, destination_spreadsheet_id, destination_sheet, new_title
        ).get("updatedSpreadsheet")
    except HttpError as e:
        print(f"ERROR {e.status_code}: {e.reason}")
        print(f'ERROR: could not copy sheet "{new_title}"')
        return

    try:
        print(
            f'SUCCESS: copied "{new_title}" sheet to {updated_spreadsheet["properties"]["title"]}'
        )
    except Exception as e:
        print("There was an error printing the sheet properties")
        print(e)
        print(f'SUCCESS: copied sheet "{new_title}')


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
