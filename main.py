from __future__ import print_function
import os.path
import pickle
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

SAMPLE_RANGE_NAME = 'Test List!A2:E246'


class GoogleSheet:
    SPREADSHEET_ID = '1z4_AXqvdIkvxWbwimm47CtTbLOKzARKL86xIDizfc2g'
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    service = None

    def __init__(self):
        creds = None
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                print('flow')
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', self.SCOPES)
                creds = flow.run_local_server(port=0)
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        self.service = build('sheets', 'v4', credentials=creds)

    def updateRangeValues(self, range, values):
        data = [{
            'range': range,
            'values': values
        }]
        body = {
            'valueInputOption': 'USER_ENTERED',
            'data': data
        }
        print(body)
        result = self.service.spreadsheets().values().batchUpdate(spreadsheetId=self.SPREADSHEET_ID,
                                                                  body=body).execute()
        print('{0} cells updated.'.format(result.get('totalUpdatedCells')))


    def printValues(self, rangee):
        # Call the Sheets API
        sheet = self.service.spreadsheets()
        result = sheet.values().get(spreadsheetId=self.SPREADSHEET_ID,
                                    range=rangee).execute()
        values = result.get('values', [])

        date = values[0][1:]
        days = 1
        for day in date:
            print(day, ': ')
            for i in range(1, len(values)):
                exl = colnum_string(days +  1)
                #print(len(values[i]), values[i])
                print(values[i][0], end='\t: ')
                if len(values[i]) > days and values[i][days] != '':
                    print(values[i][days], f'{exl}{i + 1}')
                else:
                    print(f'{exl}{i + 1}')
                   #self.updateRangeValues(range=rangee.split('!')[0]+'!'+f'{exl}{i + 1}', values=[['Пусто']]),
            print()
            days += 1
        #print('values:')
        #for i in values: print(i)

def colnum_string(n):
    string = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string = chr(65 + remainder) + string
    return string

def main():
    gs = GoogleSheet()
    test_range = 'Лист1!A1:H'
    gs.printValues(test_range)


if __name__ == '__main__':
    main()