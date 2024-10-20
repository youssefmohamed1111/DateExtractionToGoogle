import datetime
import os.path

import pandas as pd
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ["https://www.googleapis.com/auth/calendar"]


def main():
  """Shows basic usage of the Google Calendar API.
  Prints the start and name of the next 10 events on the user's calendar.
  """
  creds = None
  # The file token.json stores the user's access and refresh tokens, and is
  # created automatically when the authorization flow completes for the first
  # time.
  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  # If there are no (valid) credentials available, let the user log in.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
      flow = InstalledAppFlow.from_client_secrets_file(
          "credentials.json", SCOPES
      )
      creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  try:
    service = build("calendar", "v3", credentials=creds)
    df = pd.read_excel('Payments.xlsx')
    dates = df['Date']
    names = df['Name']
    amounts = df['Amount']
    # Call the Calendar API
    for date, name, amount in zip(dates, names,amounts):
        summary = name + " Payment: " + str(amount) + " Pounds" # Customize the event's name
        location = 'Cairo,Egypt'
        description = 'A reminder to pay: ' + str(amount) + " Pounds for " + name
        event_date = date.replace(hour=8, minute=0, second=0, microsecond=0)
        date = event_date.strftime('%Y-%m-%dT%H:%M:%S') # Date in year month day , hour, minute ,second format
        timezone = 'Africa/Cairo'
        event = {
            'summary': summary,
            'location': location,
            'description': description,
            'start': {
            'dateTime': date,
            'timeZone': timezone, #Time zone can be changed
            },
            'end': {
            'dateTime': date,
            'timeZone': timezone,
            },

            'attendees': [
            {'email': 'abcd@gmail.com'},
            ],
            'reminders': {
            'useDefault': False,
            'overrides': [
                {'method': 'email', 'minutes': 24 * 60*7}, #choose when you want to be reminded by email and popup
                {'method': 'popup', 'minutes': 24*60*7},
                {'method': 'popup', 'minutes': 10},
            ],
            },
        }


        event = service.events().insert(calendarId='primary', body=event).execute()
        print('Event created: %s' % (event.get('htmlLink')))

  except HttpError as error:
    print(f"An error occurred: {error}")


if __name__ == "__main__":
  main()