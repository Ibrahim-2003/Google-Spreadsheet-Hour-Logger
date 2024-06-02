import pandas as pd
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow,Flow
from google.auth.transport.requests import Request
import os
import pickle
import muyemail

os.system('color a')

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# here enter the id of your google sheet
SAMPLE_SPREADSHEET_ID_input = '1lwNtAjcTNYhLwQlvueP-O78ShgP_rbceCxAMvIzmDaU'
SAMPLE_RANGE_NAME = 'A1:AA1100'

def main():
    global values_input, service
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES) # here enter the name of your downloaded JSON file
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result_input = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID_input,
                                range=SAMPLE_RANGE_NAME).execute()
    values_input = result_input.get('values', [])

    if not values_input and not values_expansion:
        print('No data found.')

main()

df=pd.DataFrame(values_input[1:], columns=values_input[0])
emails = df['Email Address'].values
timestamps = df['Timestamp'].values
activities = df['Activity'].values
hours_num = df['Hour Count'].values
#hours = df[['How many hours did you spend participating/volunteering?']].values

r = open("timestamp_tracker_science.txt", "r")
tracker = r.read()
print(tracker)
tracker = tracker.split('\t')
print(tracker)

for email in range(len(emails)):
    if df['Timestamp'][email] in tracker:
        pass
    elif df['Timestamp'][email] not in tracker:
        print('\nThe email has been sent to {}'.format(emails[email]))
        send_to = emails[email]
        activity = activities[email]
        hour_count = hours_num[email]
        muyemail.smtp_live(send_to, activity, hour_count)

f = open("timestamp_tracker_science.txt", "a")

i=0
for timestamp in timestamps:
    f.write(timestamp + '\t')
    i+=1
f.close()
    

print('Success')
#print('The email has been sent to {}'.format(*emails))

input()