import logging
import os

from dotenv import load_dotenv
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

load_dotenv()

logger = logging.getLogger(__name__)
logging.basicConfig(filename='app.log', level=logging.DEBUG)


DEV_API_KEY = os.environ.get('GOOGLE_API_KEY')
SPREADSHEET_ID = os.environ.get('SPREADSHEET_ID')
ROW_OFFSET = 1
ROW_NUM_VALUES = 10


def main() -> None:
    try:
        with build('sheets', 'v4', developerKey=DEV_API_KEY) as service:
            sheet = service.spreadsheets()
            result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range='A:O').execute()
            rows = result.get('values', [])

            if len(rows) < 1:
                logger.info('No data found in spreadsheet')
                return

            new_rows = []

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
                        new_rows.append([*row[:3], *values, *row[-2:]])

                except ValueError as e:
                    logger.error('ROW_NUM:%d: Error: %s\n' % row_num, e)
    except HttpError as e:
        logger.error('%s', e)


if __name__ == '__main__':
    main()
