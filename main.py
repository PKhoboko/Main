from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from io import BytesIO
import io
from docx import Document
from flask import Flask, request, jsonify
import pickle
import os
import re
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from io import BytesIO
import os
from google_auth_oauthlib.flow import Flow, InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from google.auth.transport.requests import Request
from fastapi import FastAPI
from fastapi.responses import RedirectResponse


app = FastAPI(title="Search Google App using python")



def download_file_content(service, file_id, name):
    request = service.files().get_media(fileId=file_id)
    
    file_content = io.BytesIO()
    downloader = MediaIoBaseDownload(file_content, request)
    done = False
    while not done:
            status, done = downloader.next_chunk()
    doc = Document(io.BytesIO(file_content.getvalue()))

        # Extract and concatenate text from paragraphs in the document
    text_content = ''
    for paragraph in doc.paragraphs:
            text_content += paragraph.text + '\n'

    return text_content

def Create_Service(client_secret_file, api_name, api_version, *scopes):
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
        print(API_SERVICE_NAME, 'service created successfully')
        return service
    except Exception as e:
        print('Unable to connect.')
        print(e)
        return None
@app.get("/search-google-drive/{query}")
def Result(query):
    Client_secrete_file='client_secret_1036886342741-cs4f6svu1uoas4gt8tnfveeeajsfab18.apps.googleusercontent.com.json'
    Client_secrete_file='client_secret_1036886342741-cs4f6svu1uoas4gt8tnfveeeajsfab18.apps.googleusercontent.com.json'
    Api_name="drive"
    api_version='V3'
    scopes=['https://www.googleapis.com/auth/drive']
    service = Create_Service(Client_secrete_file, Api_name, api_version, scopes)
    response = service.files().list(q=f"fullText contains '{query}' and name contains '.docx'").execute()
    Details = []
    for file in response['files'][:50]:
        
        file_contents = download_file_content(service, file['id'], file['name'])
        Details.append({
             "Filename": file['name'], 
           "Content": file_contents[100:1000]
        })
    return jsonify(Details), 200

if __name__ == "__main__":
    app.run(debug=True)