from __future__ import print_function
import time
import json
from msrest.authentication import CognitiveServicesCredentials
from azure.cognitiveservices.vision.computervision import ComputerVisionClient
from azure.cognitiveservices.vision.computervision.models import OperationStatusCodes, VisualFeatureTypes
from googleapiclient.discovery import build
from google.oauth2 import service_account
import re

credential = json.load(open("credential.json"))
API_KEY = credential["API_KEY"]
ENDPOINT = credential["ENDPOINT"]

cv_client = ComputerVisionClient(ENDPOINT, CognitiveServicesCredentials(API_KEY))

def ocr(image_location, sheet_name):
    local_file = "./images/"+image_location+".jpg"
    response = cv_client.read_in_stream(open(local_file, 'rb'), raw=True)
    operationLocation = response.headers["Operation-Location"]
    operation_id = operationLocation.split('/')[-1]
    #print(operation_id)
    time.sleep(5)
    result = cv_client.get_read_result(operation_id)
    word_list = []

    if result.status == OperationStatusCodes.succeeded:
        read_results = result.analyze_result.read_results
        for analyzed_result in read_results:
            for line in analyzed_result.lines:
                print(line.text)
                word_list.append(line.text)
    print(word_list)

    '''Split Strings'''
    biglist = []
    spec_dict = {'A':'Body Length from HPS', 'B':'Across Shoulder', 'C':'Across Front 5" below HPS', 'D':'Across Back 5" below HPS',
                 'E':'Chest 1" below Armhole', 'F': 'Waist 19" below HPS', 'G':'Bottom Sweep', 'H':'Shirt Tail',
                 'I':'Shoulder Drop from HPS', 'J':'Sleeve Length from CB', 'K':'Sleeve Opening'}
    no_equal_list = []
    for i in word_list: #bring item in list out one by one
        if '=' in i or ':' in i:
            split_list = re.split(' =|=|= | = | :|:|: | : ', i)
            if split_list[0] in spec_dict.keys():
                split_list[0] = spec_dict[split_list[0]]
            biglist.append(split_list)
        else:
            new_list = [i]
            no_equal_list.append(new_list)

    finalList = biglist + no_equal_list #final list of what to be shown
    print("Final List =", finalList)

    '''Google Sheets'''
    '''Generate all the proper credential we need'''

    SERVICE_ACCOUNT_FILE = 'sheets_key.json'
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

    creds = None
    creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES)

    # The ID and range of a sample spreadsheet.
    SAMPLE_SPREADSHEET_ID = '1uOMifyKIO1CMSpJYdIS7Poea-oj8ORRI9Kfk6TqS8so'

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    '''
    #to read
    #result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID, range="sheet1!A1:G16").execute() #do a get request and get the information
    values = result.get('values', []) #extract the information to be in list form
    print(values) #print the value
    '''
    '''add sheet'''
    request_body = {'requests':[{'addSheet':{'properties':{'title':sheet_name}}}]} #create sheet properties
    response = sheet.batchUpdate(spreadsheetId=SAMPLE_SPREADSHEET_ID, body=request_body).execute() #create new sheet
    #aoa = [list of the data]
    #to write
    request = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=sheet_name+"!A1", valueInputOption="USER_ENTERED", body={"values":finalList}).execute()
    #for range --> just assign the starting point
    #use "USER_ENTERED" because we want the data to be treated the same way user entered in the spreadsheet.
    #if you are not sure look at these docs
    #https://developers.google.com/sheets/api/reference/rest
    #https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values#ValueRange --> for ValueRange
    #Add sheet --> https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#addsheetrequest
    #built in re library --> https://stackoverflow.com/questions/4998629/split-string-with-multiple-delimiters-in-python

ocr('shirt3', 'sheet1')
ocr('shirt4','sheet2')