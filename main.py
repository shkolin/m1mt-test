import argparse
import datetime
import logging
import os
import sys
from typing import Any

from google.auth.exceptions import MutualTLSChannelError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

logger = logging.getLogger(__name__)
logging.basicConfig(filename='app.log', level=logging.DEBUG)

ROW_OFFSET = 1
ROW_NUM_VALUES = 10
SHEET_RANGE_NAME = 'A:O'
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
]


def error(msg: str, silent: bool = False) -> None:
    logger.error(msg)
    if not silent:
        print(msg)


def create_spreadsheet(service: Any, title: str) -> Any:
    try:
        spreadsheet = service.create(
            body={'properties': {'title': title}},
            fields='spreadsheetId',
        ).execute()
        return spreadsheet['spreadsheetId']
    except HttpError as e:
        error('Failed to create spreadsheet: %s\n' % e)
        return


def get_spreadsheet_data(service: Any, spreadsheet_id: str) -> list:
    try:
        result = (
            service.values()
            .get(
                spreadsheetId=spreadsheet_id,
                range=SHEET_RANGE_NAME,
            )
            .execute()
        )
        return result.get('values', [])
    except HttpError as e:
        error('Failed to get data: %s\n' % e)
        return []


def update_spreadsheet(service: Any, spreadsheet_id: str, dataset: list) -> None:
    try:
        service.values().update(
            spreadsheetId=spreadsheet_id,
            range=SHEET_RANGE_NAME,
            valueInputOption='RAW',
            body={'values': dataset},
        ).execute()

        service.batchUpdate(
            spreadsheetId=spreadsheet_id,
            body={
                'requests': [
                    {
                        'updateSheetProperties': {
                            'properties': {'sheetId': 0, 'gridProperties': {'frozenRowCount': 1}},
                            'fields': 'gridProperties.frozenRowCount',
                        }
                    },
                    {
                        'repeatCell': {
                            'range': {'sheetId': 0, 'startRowIndex': 0, 'endRowIndex': 1},
                            'cell': {
                                'userEnteredFormat': {
                                    'textFormat': {'bold': True},
                                    'borders': {
                                        'bottom': {
                                            'style': 'SOLID_MEDIUM',
                                            'width': 3,
                                            'color': {'red': 0.6, 'green': 0.6, 'blue': 0.6},
                                        }
                                    },
                                }
                            },
                            'fields': 'userEnteredFormat.textFormat.bold',
                        }
                    },
                ]
            },
        ).execute()
    except HttpError as e:
        error('Failed to update data: %s\n' % e)


def get_credentials() -> Credentials | None:
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

    return creds


def process_dataset(rows: list[list[Any]]) -> list[list[Any]]:
    new_dataset = []

    # sheet header
    new_dataset.append(rows[0])

    for row_num, row in enumerate(rows[ROW_OFFSET:], start=2):
        try:
            values = list(map(int, row[3:-2]))
            max_num = max(values)

            if max_num < 1:
                continue

            new_values = [[0] * ROW_NUM_VALUES for _ in range(max_num)]

            for r in range(len(new_values)):
                for c, value in enumerate(values):
                    new_values[r][c] = 1 if value >= r + 1 else 0

            for values in new_values:
                new_dataset.append([*row[:3], *values, *row[-2:]])

        except ValueError as e:
            error('ROW_NUM:%d: Error: %s\n' % (row_num, e))

    return new_dataset


def main(argv: list[str]) -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument('--spreadsheet_id')
    args = parser.parse_args(argv[1:])

    if not args.spreadsheet_id:
        print('Spreadsheet ID is not provided\n')
        return

    creds = get_credentials()

    if not creds:
        print('Credentials not found')
        return

    try:
        with build('sheets', 'v4', credentials=creds) as resource:
            service = resource.spreadsheets()
            rows = get_spreadsheet_data(service, args.spreadsheet_id)

            if len(rows) < 1:
                error('No data found in spreadsheet')
                return

            new_dataset = process_dataset(rows)
            spreadsheet_id = create_spreadsheet(service, 'New Dataset %s' % datetime.datetime.now())
            update_spreadsheet(service, spreadsheet_id, new_dataset)
    except MutualTLSChannelError as e:
        error(str(e))


if __name__ == '__main__':
    main(sys.argv)
