from edapi import EdAPI
import os.path
import json
import re
import datetime


from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
SELF_REFLECTION_GRADEBOOK_ID = "1KSk6ZE21Uc3Lc5mTqD-nbrTrYRHaaZaV3BmYsbMK-Fw"
ATTENDENCE_SID = '1DyGpBhvpswW2YowTvfSWtXFQTXTOij6AOwq5d5rrVAY'
GRADEBOOK_SUBSHEET_NAME = "Grades With Only Self-Reflections"
ATTENDENCE_SUBSHEET_NAME = "Script_Edited_Attendance"

"""
Allows the user authenticate their google account, allowing the script to modify spreadsheets in their name.
Borrowed from here: https://developers.google.com/sheets/api/quickstart/python  
"""
def allow_user_to_authenticate_google_account():

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
      print("Authentication succesful")
    # Save the credentials for the next run
    with open("token.json", "w") as token:
      token.write(creds.to_json())
  return creds

def retrieve_all_threads_from_Ed():
    # initialize Ed API
    ed = EdAPI()
    # authenticate user through the ED_API_TOKEN environment variable
    ed.login()

    all_threads = []
    current_set_of_threds = []
    offset = 0

    while offset == 0 or current_set_of_threds and offset < 2000:

        current_set_of_threds = ed.list_threads(course_id = 58211, limit = 100, offset = offset)
        all_threads.extend(current_set_of_threds)
        offset += 100

    return all_threads

def generate_full_name_column(creds):
  try:
    service = build("sheets", "v4", credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = (
        sheet.values()
        .get(spreadsheetId=SELF_REFLECTION_GRADEBOOK_ID, range=GRADEBOOK_SUBSHEET_NAME)
        .execute()
    )
    first_and_last_names = result.get("values", [])

    #print(len(values), type(values))
    #print(values)
    name_combiner = lambda lst: lst[0] + " " + lst[1]
    combined_names = [name_combiner(name) for name in first_and_last_names]
    combined_names_as_string = ",".join(combined_names)

    """
    body = {
        "spreadsheetId": SPREADSHEET_ID,
        "range": '3B:B',
        "valueInputOption": 'USER_ENTERED',
        "insertDataOption": 'INSERT_ROWS',
        "resource": {
            "majorDimension": "COLUMNS",
            "values": json.dumps([combined_names])
        },
    }
    """

    request = service.spreadsheets().values().append(spreadsheetId=SELF_REFLECTION_GRADEBOOK_ID,
                                                     range='C3:C',
                                                     body={
            "majorDimension": "COLUMNS",
            "values": [combined_names]
        },
                                                     valueInputOption="USER_ENTERED"
                                                     )



    print(request.execute())




    
  except HttpError as err:
    print(err)

def retrieve_dates_to_index_mapping(sheet, sid):
    dates_result = (sheet.values().get(
    spreadsheetId=sid, 
    range=f'{GRADEBOOK_SUBSHEET_NAME}!1:1').execute()
    )
    dates = dates_result.get("values", [])[0]
    return {dates[i - 1]: i for i in range(1, len(dates) + 1)}

def retrieve_dates_to_index_mapping_attendence(sheet, sid):
    dates_result = (sheet.values().get(
    spreadsheetId=sid, 
    range='Script_Edited_Attendance!1:1').execute()
    )
    dates = dates_result.get("values", [])[0]
    return {dates[i - 1]: i for i in range(1, len(dates) + 1)}

def retrieve_names_to_index_mapping(sheet, sid):
    names_result = (sheet.values().get(
    spreadsheetId=sid, 
    range=f'{GRADEBOOK_SUBSHEET_NAME}!C:C').execute()
    )
    raw_names = names_result.get("values", [])
    cleaned_names = list(map(lambda sublist: sublist[0] if sublist else '' ,raw_names))
    #print(cleaned_names)
    return {cleaned_names[i - 1]: i for i in range(1, len(cleaned_names) + 1)}

def retrieve_names_to_index_mapping_attendence(sheet, sid):
    names_result = (sheet.values().get(
    spreadsheetId=sid, 
    range='Script_Edited_Attendance!C:C').execute()
    )
    raw_names = names_result.get("values", [])
    cleaned_names = list(map(lambda sublist: sublist[0] if sublist else '' ,raw_names))
    #print(cleaned_names)
    return {cleaned_names[i - 1]: i for i in range(1, len(cleaned_names) + 1)}

def return_notation(thread, names_to_index, dates_to_index, sheet_name=GRADEBOOK_SUBSHEET_NAME):
    print(thread['title'])
    name = thread['user']['name']
    title = thread['title']
    date = ""
    raw_date_list = re.findall("([0-9]*/[0-9]*)", title)
    # If no date was found
    if not raw_date_list:
        return f"{sheet_name}!A100"
    # A date was found                   
    date = raw_date_list[0]
    split_date = date.split("/")
    formatted_date = ""
    if not split_date:
       return f"{sheet_name}!A100"
    if len(split_date) < 2:
       return f"{sheet_name}!A100"
    if len(split_date[0]) == 1:
       formatted_date = '0' + date[0] + '/'
    else:
       formatted_date = split_date[0] + '/'
    if len(split_date[1]) == 1:
       formatted_date  += ('0' + split_date[1])
    else:
       formatted_date += split_date[1]

    row_of_date = names_to_index[name]
    if formatted_date in dates_to_index:
        col_of_date = dates_to_index[formatted_date]
    else:
        col_of_date = 1
        row_of_date = 100
    return f"{GRADEBOOK_SUBSHEET_NAME}!R" + str(row_of_date) + "C" + str(col_of_date)

def update_sheet(sheet, update_list, sid):

    batch_update_values_request_body = {
        "valueInputOption": "USER_ENTERED",
        "data": update_list
    }
    sheet.values().batchUpdate(
    spreadsheetId=sid, 
    body=batch_update_values_request_body
    ).execute()

def convert_date_to_day(date):
   date_with_year = date + '/2024'
   print(date_with_year)
   date_object = datetime.datetime.strptime(date_with_year, "%m/%d/%Y")
   return date_object.strftime("%A")


def initialize(sheet, names_to_index, dates_to_index):
    
    result = (sheet.values().get(
    spreadsheetId=SELF_REFLECTION_GRADEBOOK_ID,
    range=f'{GRADEBOOK_SUBSHEET_NAME}!C3:D38').execute()
    )
    names_and_dates = result.get("values", [])
    # Dictionary from name to list of days
    # Initialize dictionary from days to dates, {Monday: [], Tuesday: []}
    # Use prescence of / to determine whether something is a date.
    # for date in dates, add 
    # for name in names_to_index, add {'range': notation, 'values': [['No']]} to update list
    name_to_list_of_days = { record[0]:record[1].strip().split(',') for record in names_and_dates}
    days_to_date_index = {'Monday':[], 'Tuesday':[], 'Wednesday':[], 'Thursday':[], 'Friday':[]}
    
    for potential_date, index in dates_to_index.items():
       if '/' in potential_date:
          day = convert_date_to_day(potential_date)
          days_to_date_index[day].append(index)
    update_list = []
    for name, days in name_to_list_of_days.items():
       for day in days:
            day = day.strip()
            date_index_list = days_to_date_index[day]
            for date_index in date_index_list:               
                notation = f"{GRADEBOOK_SUBSHEET_NAME}!R" + str(names_to_index[name]) + "C" + str(date_index)
                update_list.append({'range': notation, 'values': [['No']]})
    update_sheet(sheet, update_list, SELF_REFLECTION_GRADEBOOK_ID)
    
       

def perform_specified_task(task):
    creds = allow_user_to_authenticate_google_account()
    threads = retrieve_all_threads_from_Ed()
    # Retrieve dates row from sheet
    # Make that list a dictionary

    service = build("sheets", "v4", credentials=creds)
    

    # Call the Sheets API
    sheet = service.spreadsheets()


    names_to_index = retrieve_names_to_index_mapping(sheet, SELF_REFLECTION_GRADEBOOK_ID)
    dates_to_index = retrieve_dates_to_index_mapping(sheet, SELF_REFLECTION_GRADEBOOK_ID)
    update_list = []
    if task == "Self Reflections":
        for thread in threads:
            if thread['category'] == "Self Reflections":
                notation = return_notation(thread, names_to_index, dates_to_index)
                update_list.append({'range': notation, 'values': [['Yes']]})
        
        update_sheet(sheet, update_list, SELF_REFLECTION_GRADEBOOK_ID)
    elif task == "Lecture Makeups":
        for thread in threads:
            if thread['category'] == "Self Reflections":
                notation = return_notation(thread, names_to_index, dates_to_index)
                update_list.append({'range': notation, 'values': [['Yes']]})
    elif task == "Initialize":
       initialize(sheet, names_to_index, dates_to_index)
    elif task == "Attendence":
        attendence()
       
def attendence():
   service = build("sheets", "v4", credentials=allow_user_to_authenticate_google_account())
   sheet = service.spreadsheets()
   threads = retrieve_all_threads_from_Ed()
   names_to_index = retrieve_names_to_index_mapping_attendence(sheet, ATTENDENCE_SID)
   dates_to_index = retrieve_dates_to_index_mapping_attendence(sheet, ATTENDENCE_SID)
   print(dates_to_index)
   update_list = []
   for thread in threads:
      if thread['category'] == "Lecture makeup":
            notation = return_notation(thread, names_to_index, dates_to_index, sheet_name=ATTENDENCE_SUBSHEET_NAME)
            update_list.append({'range': notation, 'values': [['TRUE']]})
   update_sheet(sheet, update_list, ATTENDENCE_SID)


def surveys():
   service = build("sheets", "v4", credentials=allow_user_to_authenticate_google_account())
   sheet = service.spreadsheets()
   threads = retrieve_all_threads_from_Ed()
   names_to_index = retrieve_names_to_index_mapping_attendence(sheet, ATTENDENCE_SID)
   names = set(names_to_index.keys())
   update_list = []
   for thread in threads:
      if thread['category'] == "Survey":
            if thread['user']['name'] in names:
                names.remove(thread['user']['name'])


def main():
    perform_specified_task("Self Reflections")
    attendence()

main()
