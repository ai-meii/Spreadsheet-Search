
# This part is the imports, permissions and set-up vvv

import os
from google.oauth2 import service_account, credentials
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly', 
'https://www.googleapis.com/auth/drive.metadata.readonly']

service_acc_file = '...' #Obtain a service account file from Google Cloud Console

credentials = service_account.Credentials.from_service_account_file(service_acc_file, scopes=SCOPES)

# Google drive api set up 

drive_service = build('drive', 'v3', credentials=credentials)

folder_id = '...' #Can be found in the URL of the folder, probably afer /folder/ or id=

query = f"'{folder_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"

response = drive_service.files().list(q=query,fields='files(id, name)').execute()

items = response.get('files', [])

if not items:
    print('No Google Sheets found in this folder.')
else:
    fileList = []
    for item in items:
        fileList.append(item['id'])

service = build('sheets', 'v4', credentials=credentials)

sheet = service.spreadsheets()

# Gets user input and edits the format if necessary vvv

def getInput():
    pc_number = input("Enter the PC Number: ")
    return pc_number

def inputEdit(pc_number):
    if "PC" in pc_number:
        PC = pc_number[0:2]
        if " " not in pc_number:
            pc_number = PC + " " + pc_number[2:]
        if PC.isupper() == False:
            pc_number = (PC.upper()) + pc_number[2:]
    elif "PC" not in pc_number:
        pc_number = "PC " + pc_number
    return(pc_number)

# Parsing through spreadsheets in the folder vvv

def search(user_input):
    for id in fileList:
        tabs_read = service.spreadsheets().get(spreadsheetId = id, fields='sheets.properties.title').execute()
        tabs = [sheet['properties']['title'] for sheet in tabs_read['sheets']]
        for tabName in tabs:
            range = f'{tabName}!A:A'
            sheet_read = sheet.values().get(spreadsheetId = id, range=range).execute()
            values = sheet_read.get('values', [])
            flat_values = [cell for index in values for cell in index]
            if user_input in flat_values:
                return tabName
            else:
                continue

user_input = inputEdit(pc_number=getInput())

print(search(user_input))
