from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from os import path


import argparse
parser = argparse.ArgumentParser(description='Setup credentials and tokens for Google API')
parser.add_argument("-drive",action="store_true",default=False,dest="drive",help="Setup for Drive API")
parser.add_argument("-docs",action="store_true",default=False,dest="docs",help="Setup for Docs API")
parser.add_argument("-sheets",action="store_true",default=False,dest="sheets",help="Setup for Sheet API")

CURDIR = path.dirname(path.abspath(__file__))
SCOPES_DOCS = ['https://www.googleapis.com/auth/documents']
SCOPES_DRIVE = ['https://www.googleapis.com/auth/drive']
SCOPES_SHEETS = ['https://www.googleapis.com/auth/spreadsheets']

def generate_token_docs():

    creds = None
    if os.path.exists(path.join(CURDIR,"tokens",'token_docs.pickle')):
        with open(path.join(CURDIR,"tokens",'token_docs.pickle'), 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                path.join(CURDIR,'credentials','credentials_docs.json'), SCOPES_DOCS)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(path.join(CURDIR,"tokens",'token_docs.pickle'), 'wb') as token:
            pickle.dump(creds, token)

    service = build('docs', 'v1', credentials=creds)

def generate_token_sheets():

    creds = None
    if os.path.exists(path.join(CURDIR,"tokens",'token_sheets.pickle')):
        with open(path.join(CURDIR,"tokens",'token_sheets.pickle'), 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                path.join(CURDIR,'credentials','credentials_sheets.json'), SCOPES_SHEETS)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(path.join(CURDIR,"tokens",'token_sheets.pickle'), 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

def generate_token_drive():

    creds = None
    if os.path.exists(path.join(CURDIR,"tokens",'token_drive.pickle')):
        with open(path.join(CURDIR,"tokens",'token_drive.pickle'), 'rb') as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                path.join(CURDIR,'credentials','credentials_drive.json'), SCOPES_DRIVE)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(path.join(CURDIR,"tokens",'token_drive.pickle'), 'wb') as token:
            pickle.dump(creds, token)



if __name__ == '__main__':
    args = parser.parse_args()

    if args.drive:
        generate_token_drive()
    if args.sheets:
        generate_token_sheets()
    if args.docs:
        generate_token_docs()

