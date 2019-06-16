import datetime
import os.path
import pickle
from itertools import zip_longest

from google.auth.transport.requests import Request
# from google_api_python_client.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


def convert_xls_datetime(xls_date):
    return (datetime.datetime(1899, 12, 30)
            + datetime.timedelta(days=xls_date))


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
MORTGAGE_SPREADSHEET_ID = '114cw7xViCFS0tqvJ_Km3WhgN9Ot6PGBorLqYAlVskp8'
PAYMENTS_RANGE_NAME = "'Sheet1'!E4:E"
DATES_RANGE_NAME = "'Sheet1'!A4:A"


def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    range_names = [DATES_RANGE_NAME, PAYMENTS_RANGE_NAME]
    # result = sheet.values().get(spreadsheetId=MORTGAGE_SPREADSHEET_ID,
    #                             range=PAYMENTS_RANGE_NAME).execute()
    results = sheet.values().batchGet(spreadsheetId=MORTGAGE_SPREADSHEET_ID, ranges=range_names,
                                      valueRenderOption='UNFORMATTED_VALUE').execute()
    ranges = results.get('valueRanges', [])
    dates = ranges[0].get('values', [])
    payments = ranges[1].get('values', [])

    if not dates or not payments:
        print('No data found.')
    else:
        assert payments[0][0] == "payment"
        assert dates[0][0] == "dates"
        # print('Name, Major:')
        # for row in values:
        #     # Print columns A and E, which correspond to indices 0 and 4.
        #     print('%s, %s' % (row[0], row[4]))
        for date, payment in zip_longest(dates[1:], payments[1:]):
            if payment:
                payment = payment[0]
            date = date[0]
            if convert_xls_datetime(date) >= datetime.datetime.today() and not payment:
                try:
                    print("making an append")
                    payments.append([3000])
                    result = sheet.values().update(spreadsheetId=MORTGAGE_SPREADSHEET_ID, range=PAYMENTS_RANGE_NAME,
                                                   valueInputOption='RAW', body={'values': payments}).execute()
                    # result = sheet.values().append(spreadsheetId=MORTGAGE_SPREADSHEET_ID, range=PAYMENTS_RANGE_NAME, valueInputOption='RAW',responseValueRenderOption='UNFORMATTED_VALUE', body={'values':[[3000]]}).execute()
                    break
                except Exception as e:
                    raise (e)

        print("done looking to append")


if __name__ == '__main__':
    main()
