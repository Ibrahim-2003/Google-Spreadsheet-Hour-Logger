import pandas as pd
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow,Flow
from google.auth.transport.requests import Request
import os
import pickle
import pyperclip


os.system('color a')

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# here enter the id of your google sheet
SAMPLE_SPREADSHEET_ID_input = '1_Wh40HAmArHjkzvtvyKvPHVFAERq4FMZpNw7Z2ToLLU'
SAMPLE_RANGE_NAME = 'A1:AA1100'

def main(RANGE_NAME):
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
                'D:\Desktop Backup\School\Homework\Statistics\Python Stuff\credentials.json', SCOPES) # here enter the name of your downloaded JSON file
            creds = flow.run_local_server(port=0)
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result_input = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID_input,
                                range=RANGE_NAME).execute()
    values_input = result_input.get('values', [])

    if not values_input and not values_expansion:
        print('No data found.')


completed = []
likely = []
unlikely = []
completed_jr = []
likely_jr = []
unlikely_jr = []
completed_sr = []
likely_sr = []
unlikely_sr = []
grade_levels = ['Seniors', 'Juniors', 'Sophomores']


i=0
for grade in grade_levels:
    ranges = ['A1:C20', 'A21:C28', 'A29:C32']
    main(ranges[i])
    df=pd.DataFrame(values_input[1:], columns=values_input[0])
    students = df['Students'].values
    hours_left = df['LEFT'].values
    #hours = df[['How many hours did you spend participating/volunteering?']].values
    for left in range(len(hours_left)):
        if grade == grade_levels[0]:
            if float(hours_left[left]) <= 0:
                completed.append(students[left])
            elif float(hours_left[left]) > 0 and float(hours_left[left]) <= 5:
                likely.append(students[left])
            elif float(hours_left[left]) > 5:
                unlikely.append(students[left])
            results = f'{grade.upper()}:\nStudents with Hours Completed: \n{completed}\n\n---------------------------\n\nStudents within 5 Hours Left: \n{likely}\n\n---------------------------\n\nStudents with more than 5 Hours Left: \n{unlikely}\n\n---------------------------\n\n'
            print(results)
        elif grade == grade_levels[1]:
            if float(hours_left[left]) <= 0:
                completed_jr.append(students[left])
            elif float(hours_left[left]) > 0 and float(hours_left[left]) <= 5:
                likely_jr.append(students[left])
            elif float(hours_left[left]) > 5:
                unlikely_jr.append(students[left])
            jr_results = f'{grade.upper()}:\nStudents with Hours Completed: \n{completed_jr}\n\n---------------------------\n\nStudents within 5 Hours Left: \n{likely_jr}\n\n---------------------------\n\nStudents with more than 5 Hours Left: \n{unlikely_jr}\n\n---------------------------\n\n'
            print(jr_results)
        else:
            if float(hours_left[left]) <= 0:
                completed_sr.append(students[left])
            elif float(hours_left[left]) > 0 and float(hours_left[left]) <= 5:
                likely_sr.append(students[left])
            elif float(hours_left[left]) > 5:
                unlikely_sr.append(students[left])
            sr_results = f'{grade.upper()}:\nStudents with Hours Completed: \n{completed_sr}\n\n---------------------------\n\nStudents within 5 Hours Left: \n{likely_sr}\n\n---------------------------\n\nStudents with more than 5 Hours Left: \n{unlikely_sr}\n\n---------------------------\n\n'
            print(sr_results)
    print('\n\n\n')
    i+=1    
final_result = results + '\n\n\n' + jr_results + '\n\n\n' + sr_results
pyperclip.copy(final_result)

input()