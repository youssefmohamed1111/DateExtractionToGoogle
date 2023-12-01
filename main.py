import googleapiclient
import pandas as pd
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

df = pd.read_excel('Payments.xlsx')
dates = df['Date']
names = df['Name']
# print(dates,names)

Scopes = ['https://www.googleapis.com/auth/calendar']
account = 'token.json'
creds = Credentials.from_authorized_user_file("token.json", Scopes)

service = googleapiclient.discovery.build('calendar', 'v3', credentials=creds)
for date, name in zip(dates, names):
  summary = name + " Payment"
  event_date = date.replace(hour=8, minute=0, second=0, microsecond=0)

  Date = event_date.strftime('%Y-%m-%dT%H:%M:%S')
  event = {
    'summary': summary,
    'location': 'Cairo,Egypt',
    'description': 'A reminder to pay',
    'start': {
      'dateTime': Date,
      'timeZone': 'Africa/Cairo',
    },
    'end': {
      'dateTime': Date,
      'timeZone': 'Africa/Cairo',
    },

    'attendees': [
      {'email': 'bbtengan4@gmail.com'},

    ],
    'reminders': {
      'useDefault': False,
      'overrides': [
        {'method': 'email', 'minutes': 24 * 60},
        {'method': 'popup', 'minutes': 10},
      ],
    },
  }

  event = service.events().insert(calendarId='primary', body=event).execute()
  print('Event created: %s' % (event.get('htmlLink')))
