
import os
import sys
import pickle
import datetime
import pandas as pd
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import Flow, InstalledAppFlow
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from pyzillow.pyzillow import ZillowWrapper, GetDeepSearchResults, GetUpdatedPropertyDetails


class ZillowSpreadsheetDemo:

    """
    Blueprint of Google Spreadsheet
    """

    # Constructor function
    def __init__(self):
        super(ZillowSpreadsheetDemo, self).__init__()
        self.response_values_list = None
        self.spreadsheet_ID = None
        self.service = None

    def create_spreadsheet(self):

        # Assemble Google API service
        CSJ_FILE = 'client_secret.json'
        API_NAME = 'sheets'
        API_VERS = 'v4'
        SVC_SCPE = ['https://www.googleapis.com/auth/spreadsheets']

        self.service = self.create_service(CSJ_FILE, API_NAME, API_VERS, SVC_SCPE)
        self.spreadsheet_ID = '1VBzntQBEW-pPRtgl07iQCoY9RXvI6sbL5nQDH3B-714'

        response = self.service.spreadsheets().values().get(
            spreadsheetId = self.spreadsheet_ID,
            majorDimension = 'ROWS',
            range = 'Sheet1!B3:C5'
        ).execute()

        self.response_values_list = (response['values'])

    def update_spreadsheet(self):

        insert_values = []
        for zip_code, add_ress in self.response_values_list:
            insert_values.append(self.connect_to_zillow(add_ress, zip_code))

        worksheet_name = 'Sheet1!'
        cell_range_insert = 'D3'
        value_range_body = {
            'majorDimension': 'ROWS',
            'values': insert_values
        }
        self.service.spreadsheets().values().append(
            spreadsheetId = self.spreadsheet_ID,
            valueInputOption = 'USER_ENTERED',
            range = worksheet_name + cell_range_insert,
            body = value_range_body
        ).execute()

    def connect_to_zillow(self, address, zipcode):

        zillow_data = ZillowWrapper('X1-ZWz18nuirxd4i3_2avji')
        deep_search = zillow_data.get_deep_search_results(address, zipcode, False)
        result = GetDeepSearchResults(deep_search)

        # print('\t: {0} : {1} : {2}'.format(result.home_size, result.year_built, result.zestimate_amount))
        return (result.zestimate_amount, result.home_size, result.year_built)

    def create_service(self, client_secret_file, api_name, api_version, *scopes):
        print(client_secret_file, api_name, api_version, scopes, sep='-')
        CLIENT_SECRET_FILE = client_secret_file
        API_SERVICE_NAME = api_name
        API_VERSION = api_version
        SCOPES = [scope for scope in scopes[0]]
        print(SCOPES)
    
        cred = None
    
        pickle_file = f'token_{API_SERVICE_NAME}_{API_VERSION}.pickle'
        # print(pickle_file)
    
        if os.path.exists(pickle_file):
            with open(pickle_file, 'rb') as token:
                cred = pickle.load(token)
    
        if not cred or not cred.valid:
            if cred and cred.expired and cred.refresh_token:
                cred.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
                cred = flow.run_local_server()
    
            with open(pickle_file, 'wb') as token:
                pickle.dump(cred, token)
    
        try:
            service = build(API_SERVICE_NAME, API_VERSION, credentials=cred)
            print('\nCreated Service Successfully\n')
            return service
        except Exception as e:
            print('Unable to connect.')
            print(e)
            return None
    
    def convert_to_RFC_datetime(self, year=1900, month=1, day=1, hour=0, minute=0):
        dt = datetime.datetime(year, month, day, hour, minute, 0).isoformat() + 'Z'
        return dt


if __name__ == '__main__':
    app = ZillowSpreadsheetDemo()
    app.create_spreadsheet()
    app.update_spreadsheet()
