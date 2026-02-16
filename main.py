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


def main() -> None:
    try:
        with build('sheets', 'v4', developerKey=DEV_API_KEY) as service:
            sheet = service.spreadsheets()
            result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range='A:O').execute()
            values = result.get('values', [])

            if len(values) > 1:
                for row_num, row in enumerate(values[ROW_OFFSET:], start=2):
                    try:
                        date, regioan, city = row[:3]
                        long, lat = row[-2:]
                        row_num_values = list(map(int, row[3:-2]))
                        row_max_num_value = max(row_num_values)
                    except ValueError as e:
                        msg = 'ROW_NUM:%d: Error: %s\n' % (row_num, e)
                        logger.error(msg)
                        print(msg)
    except HttpError as e:
        logger.error('%s', e)
        print(str(e))


if __name__ == '__main__':
    main()
