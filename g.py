import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
import csv

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = "1lIAdR685J_SS2Mz2KYtjUs8EK70Wg_6scVpmGWIY5cE"


def update_spreadsheet_from_csv(spreadsheet_id=SAMPLE_SPREADSHEET_ID, range_name="market_ocr!A1", csv_file_path="output.csv", credentials_file="credentials.json", token_file="token.json"):
    """Updates a Google Sheets spreadsheet with data from a CSV file."""
    creds = None
    if os.path.exists(token_file):
        creds = Credentials.from_authorized_user_file(token_file, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                credentials_file, SCOPES
            )
            creds = flow.run_local_server(port=0)
        with open(token_file, "w") as token:
            token.write(creds.to_json())

    try:
        service = build("sheets", "v4", credentials=creds)
        sheet = service.spreadsheets()

        with open(csv_file_path, mode="r", encoding="utf-8") as file:
            csv_reader = csv.reader(file)
            data = list(csv_reader)

        body = {
            "values": data
        }

        # Clear the existing data in the spreadsheet
        sheet.values().clear(
            spreadsheetId=spreadsheet_id,
            range="market_ocr!A1:C",
            body={}
        ).execute()

        result = sheet.values().update(
            spreadsheetId=spreadsheet_id,
            range=range_name,
            valueInputOption="RAW",
            body=body
        ).execute()

        print(f"{result.get('updatedCells')} cells updated.")
    except HttpError as err:
        print(err)
