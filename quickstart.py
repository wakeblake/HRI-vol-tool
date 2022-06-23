#!/usr/bin/env python
# coding: utf-8

# See https://developers.google.com/apps-script/api/quickstart/python#further_reading
# See https://github.com/PyGithub/PyGithub#pygithub

def oauth2_google(project_creds_json): 
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

    #TODO: ADD TRY/EXCEPT HERE TO DELETE STALE TOKEN?
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                project_creds_json, SCOPES)
            creds = flow.run_local_server(port=8080)
            
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    return creds


def get_script(creds, scriptId):
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

    
def save_script(response, dir_path):
    """Saves script object response as files to local directory"""
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

    
def scripts_from_github(repo, file_or_dir):
    """Pulls scripts from Github repo ("owner/project")"""
    from github import Github
    import os
    import json
    
    home = os.getenv('HOME')
    fname = {'json':'JSON', 'gs':'SERVER_JS', 'html':'HTML'}
    
    # token = os.getenv('GITHUB_TOKEN')
    # Build github api request from creds stored in gh package
    token_path = home + '/.config/gh/hosts.yml'
    with open(token_path) as f:
        for l in f:
            _l = l.strip()
            if _l.startswith('oauth_token'):
                token = _l.split(': ')[1]

    g = Github(token)
    repo = g.get_repo(repo)
    files = repo.get_contents(file_or_dir)
    files = list(
                map(
                    lambda f: {
                        "name":f.name.split('.')[0], 
                        "source":f.decoded_content.decode(), 
                        "type":fname[f.name.split('.')[1]]
                    }, 
                    files
                )
            )
    return files


def create_script(creds, files, title):
    """Calls the Apps Script API and creates container-bound script"""
    from googleapiclient import errors
    from googleapiclient.discovery import build
    import json
    
    # Build Google sheets request
    spreadsheet = {
        'properties': {
            'title': title
        }
    }
    
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
    
    # Build Google scripts request
    script = {
        'title': title,
        'parentId': spreadsheetId
    }

    service = build('script', 'v1', credentials=creds)
    
    try:  
        # Create script
        # Must set up installable triggers separately
        script = service.projects().create(body=script).execute()
        scriptId = script.get('scriptId')
        print('Script Id: {}'.format(scriptId))
        script = service.projects().updateContent(scriptId=scriptId, body={'files': files, 'scriptId':scriptId}).execute()
        scriptParentId = service.projects().get(scriptId=scriptId, fields='parentId').execute()
        if spreadsheetId == scriptParentId['parentId']:
            print('Successfully created containter-bound script')
        service.close()
    except errors.HttpError as error:
        print(error.content)
    

