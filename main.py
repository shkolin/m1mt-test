import argparse
import datetime
import logging
import os
import sys
from typing import Any

from arcgis import GIS
from dotenv import load_dotenv
from google.auth.exceptions import MutualTLSChannelError
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

load_dotenv()
logger = logging.getLogger(__name__)
logging.basicConfig(filename='app.log', level=logging.DEBUG)

ROW_OFFSET = 1
ROW_NUM_VALUES = 10
SHEET_RANGE_NAME = 'A:O'
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
]
GIS_FEATURE_LAYER_ID = os.environ.get('GIS_FEATURE_LAYER_ID')
GIS_API_TOKEN = os.environ.get('GIS_API_TOKEN')


def error(msg: str, silent: bool = False) -> None:
    logger.error(msg)
    if not silent:
        print(msg)


def to_float(value: str) -> float | None:
    try:
        return float(value.replace(',', '.'))
    except ValueError:
        return None


def create_spreadsheet(service: Any, title: str) -> str | None:
    try:
        spreadsheet = service.create(
            body={'properties': {'title': title}},
            fields='spreadsheetId',
        ).execute()
        return spreadsheet['spreadsheetId']
    except HttpError as e:
        error('Failed to create spreadsheet: %s' % e)
        return None


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
        error('Failed to get data: %s' % e)
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
    for i, row in enumerate(rows[ROW_OFFSET:], start=2):
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
                new_dataset.append([*row[:3], *values, *list(map(to_float, row[-2:]))])
        except ValueError as e:
            error('ROW_NUM:%d: Error: %s\n' % (i, e))

    return new_dataset


def export_dataset_to_arcgis(data: list[list[Any]]) -> None:
    gis = GIS('https://www.arcgis.com', token=GIS_API_TOKEN)
    item = gis.content.get(GIS_FEATURE_LAYER_ID)
    layer = item.layers[0]

    attributes = [
        'date',
        'region',
        'city',
        'value_1',
        'value_2',
        'value_3',
        'value_4',
        'value_5',
        'value_6',
        'value_7',
        'value_8',
        'value_9',
        'value_10',
        'long',
        'lat',
    ]

    features = []
    for row in data:
        long, lat = row[-2:]
        feature = {
            'attributes': {attributes[i]: value for i, value in enumerate(row)},
            'geometry': {
                'x': long,
                'y': lat,
                'spatialReference': {'wkid': 4326},
            },
        }
        features.append(feature)

    layer.edit_features(adds=features)


def main(argv: list[str]) -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument('--spreadsheet_id', required=True)

    args = parser.parse_args(argv[1:])

    creds = get_credentials()

    if not creds:
        error('Credentials not found')
        return None

    try:
        with build('sheets', 'v4', credentials=creds) as resource:
            service = resource.spreadsheets()
            rows = get_spreadsheet_data(service, args.spreadsheet_id)

            if len(rows) < 1:
                error('No data found in spreadsheet')
                return None

            new_dataset = process_dataset(rows)
            new_dataset.insert(0, rows[0])
            spreadsheet_id = create_spreadsheet(service, 'New Dataset %s' % datetime.datetime.now())
            if spreadsheet_id:
                update_spreadsheet(service, spreadsheet_id, new_dataset)

        export_dataset_to_arcgis(new_dataset[1:])
    except MutualTLSChannelError as e:
        error(str(e))


if __name__ == '__main__':
    main(sys.argv)
