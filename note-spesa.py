from __future__ import print_function
from datetime import date
import os.path
from wsgiref.util import shift_path_info
import holidays
from random import Random
from calendar import Calendar
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SPREADSHEET_ID = '1_G86nEoEzipwL5Xiw8JVqEGXXNW964AZowD0o5-Ttiw'
#The 'writable' range in the sheet
WRITABLE_RANGE = 'A17:G44'
TOTAL_DISTANCE = 21661
YEAR = 2020

DESTINATIONS = [
    {
        'name': 'Italmondo MI',
        'city': 'Milano',
        'distance': 512,
        'weight': 1
    },
    {
        'name': 'Italmondo PD',
        'city': 'Vigonza',
        'distance': 16,
        'weight': 9
    },
    {
        'name': 'Bierre PD',
        'city': 'Camposampiero',
        'distance': 50,
        'weight': 6
    },
    {
        'name': 'Arca VR',
        'city': 'Verona',
        'distance': 194,
        'weight': 3
    },
    {
        'name': 'Service 2000 TV',
        'city': 'Casale sul Sile',
        'distance': 110,
        'weight': 6
    },
    {
        'name': 'Arca PD',
        'city': 'Padova',
        'distance': 18,
        'weight': 8
    },
    {
        'name': 'Trilem PD',
        'city': 'Grantorto',
        'distance': 80,
        'weight': 7
    },
    {
        'name': 'Geo Logistics SRL',
        'city': 'Ceresone',
        'distance': 42,
        'weight': 6
    },
    {
        'name': 'Campagnolo Trasporti',
        'city': 'Tezze sul Brenta',
        'distance': 91,
        'weight': 6
    },
    {
        'name': 'Arcese TV',
        'city': 'Treviso',
        'distance': 106,
        'weight': 4
    },
    {
        'name': 'Torello AV',
        'city': 'Avellino',
        'distance': 1458,
        'weight': 1
    },
    {
        'name': 'Torello PD',
        'city': 'Padova',
        'distance': 12,
        'weight': 8
    },
    {
        'name': 'Susa PD',
        'city': 'Padova',
        'distance': 5,
        'weight': 5
    },
    {
        'name': 'Susa PG',
        'city': 'Perugia',
        'distance': 702,
        'weight': 1
    }
]

def auth():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    return creds

def clear(sheet):
    sheet.values().batchClear(
        spreadsheetId=SPREADSHEET_ID,
        body={
            'ranges': [ WRITABLE_RANGE ]
        }
    ).execute()

def gen_data():
    
    avg_distance = 0
    random = Random()
    pointers = []
    data = []
    for month in range(12):
        data.append([])

    for destination in DESTINATIONS:
        avg_distance += destination['distance']
        for _ in range(destination['weight']):
            pointers.append(destination)

    avg_distance /= len(DESTINATIONS)

    for month in range(1, 13):
        monthly_distance = 0
        month_dates = workdates(YEAR, month)
        while monthly_distance < (TOTAL_DISTANCE / 12 - avg_distance):
            destination = random.choice(pointers)
            travel = []
            travel.append((random.choice(month_dates)).isoformat())
            travel.append(destination['name'])
            travel.append('Padova')
            travel.append(destination['city'])
            travel.append(destination['city'])
            travel.append('Padova')
            travel.append(destination['distance'])
            data[month - 1].append(travel)
            monthly_distance += destination['distance']

    return data

def workdates(year, month):
    dates = []
    it_holidays = holidays.IT()
    calendar = Calendar()
    for date in calendar.itermonthdates(year, month):
        dates.append(date)
    workdates = filter(
        lambda date: 
            date.month == month 
            and date.isoweekday() != 6
            and date.isoweekday() != 7
            and date not in it_holidays,
        dates
    )
    
    result = []
    for date in workdates:
        result.append(date)

    return result

def write(sheet, data):
    sheet.values().update(spreadsheetId=SPREADSHEET_ID, range=WRITABLE_RANGE, body={ 'values': data }, valueInputOption='USER_ENTERED').execute()

def main():
    creds = auth()

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    clear(sheet)
    write(sheet, gen_data()[0])

if __name__ == '__main__':
    main()