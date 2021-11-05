#!/usr/bin/env python
# coding: utf-8

# In[21]:


def oauthGoogle(scriptId):
    
    """Calls the Apps Script API."""
    
    from googleapiclient import errors
    from googleapiclient.discovery import build
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    import os.path

    SCOPES = ['https://www.googleapis.com/auth/script.projects']

    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
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

    # Build request
    service = build('script', 'v1', credentials=creds)
    
    # Call the Apps Script API
    try:
        # Get project
        response = service.projects().getContent(scriptId=scriptId).execute()
        service.close()
        print('Response from Apps Script API received')
        return response
    except errors.HttpError as error:
        # The API encountered a problem.
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

