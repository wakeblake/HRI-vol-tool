#!/usr/bin/env python
# coding: utf-8

# In[51]:


def oauth2Google(): 
    """Creates authorization credentials for downstream functions"""
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    import os.path

    SCOPES = [
        'https://www.googleapis.com/auth/script.projects',
        'https://www.googleapis.com/auth/spreadsheets'
    ]
    
    creds = None
    
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=8080)
            
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    return creds


def getScript(creds, scriptId):
    """Calls the Apps Script API and returns script object response """
    from googleapiclient import errors
    from googleapiclient.discovery import build
    
    # Build request
    service = build('script', 'v1', credentials=creds)
    
    try:
        # Get project
        response = service.projects().getContent(scriptId=scriptId).execute()
        service.close()
        print('Response from Apps Script API received')
        return response
    except errors.HttpError as error:
        print(error.content)

    
def saveScript(response, dir_path):
    from os import getcwd as cwd
    from os import chdir
    import sys 

    curdir = cwd()
    chdir(dir_path)
    
    # Write files to dir
    origin = sys.stdout

    for f in response['files']:
        if f['type'] == 'JSON':
            ftype = '.json'
        elif f['type'] == 'SERVER_JS':
            ftype = '.gs'
        elif f['type'] == 'HTML':
            ftype = '.html'

        fname = '{}{}'.format(f['name'], ftype)

        with open(fname, 'w+') as file:
            sys.stdout = file
            print(f['source'])
            sys.stdout = origin
        
        print('Saved {} to {}'.format(fname, cwd()))
    
    chdir(curdir)
    

def createScript(creds, title):
    """Calls the Apps Script API and creates container-bound script"""
    from googleapiclient import errors
    from googleapiclient.discovery import build
    import json

    spreadsheet = {
        'properties': {
            'title': title
        }
    }
    
    # Build request
    service = build('sheets', 'v4', credentials=creds)
    
    try:
        # Create  spreadsheet
        spreadsheet = service.spreadsheets().create(body=spreadsheet, fields='spreadsheetId,spreadsheetUrl').execute()
        spreadsheetId = spreadsheet.get('spreadsheetId')
        spreadsheetUrl = spreadsheet.get('spreadsheetUrl')
        print('Spreadsheet Id: {}'.format(spreadsheetId))
        print('Spreadsheet URL: {}'.format(spreadsheetUrl))
        service.close()
    except errors.HttpError as error:
        print(error.content)
    
    script = {
        'title': title,
        'parentId': spreadsheetId
    }

    # Build request
    service = build('script', 'v1', credentials=creds)
    
    try:  
        # Create script
        response = service.projects().create(body=script).execute()
        print('Script Id: {}'.format(response.get('scriptId')))
        if spreadsheetId == response.get('parentId'):
            print('Successfully created containter-bound script')
        service.close()
    except errors.HttpError as error:
        print(error.content)
    


# In[20]:


creds = oauth2Google()


# In[52]:


createScript(creds, 'Test CB Script')


# In[ ]:




