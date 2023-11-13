from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
from io import BytesIO
from docx import Document
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import flask
from flask import session, redirect, url_for, jsonify

app = flask.Flask(__name__)
app.secret_key = 'your_generated_secret_key'

# Google Drive API settings
CLIENT_SECRET_FILE = 'client_secret_1036886342741-cs4f6svu1uoas4gt8tnfveeeajsfab18.apps.googleusercontent.com.json'
API_NAME = 'drive'
API_VERSION = 'v3'
SCOPES = ['https://www.googleapis.com/auth/drive']

def authorize_user():
    flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
    authorization_url, _ = flow.authorization_url(prompt='consent')

    print(f'Please go to this URL to authorize: {authorization_url}')
    authorization_code = input('Enter the authorization code: ')

    credentials = flow.fetch_token(code=authorization_code)

    return credentials

def create_drive_service(credentials):
    service = build(API_NAME, API_VERSION, credentials=credentials)
    return service

def download_file_content(service, file_id, name):
    request = service.files().export_media(fileId=file_id, mimeType='application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    
    file_content = BytesIO()
    downloader = MediaIoBaseDownload(file_content, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    
    doc = Document(BytesIO(file_content.getvalue()))

    text_content = ''
    for paragraph in doc.paragraphs:
        text_content += paragraph.text + '\n'

    return text_content

def setup_headless_browser():
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    return webdriver.Chrome(options=chrome_options)

@app.route('/')
def index():
    if 'credentials' not in session:
        return redirect(url_for('login'))

    credentials = session['credentials']
    service = create_drive_service(credentials)

    browser = setup_headless_browser()
    browser.get('https://drive.google.com')
    
    # You can add further Selenium actions here to interact with the Google Drive web interface.
    # For example, you can find elements, click buttons, and navigate through folders.

    # Capture a screenshot (you can customize this part)
    browser.save_screenshot('/tmp/screenshot.png')

    browser.quit()

    response = service.files().list(q="mimeType='application/vnd.google-apps.document'").execute()
    details = []

    for file in response.get('files', []):
        file_contents = download_file_content(service, file['id'], file['name'])
        details.append({
            "Filename": file['name'], 
            "Content": file_contents[100:1000]
        })

    return jsonify(details), 200

@app.route('/login')
def login():
    credentials = authorize_user()
    session['credentials'] = credentials_to_dict(credentials)
    return redirect(url_for('index'))

@app.route('/logout')
def logout():
    if 'credentials' in session:
        del session['credentials']
    return redirect(url_for('index'))

def credentials_to_dict(credentials):
    return {'token': credentials.token,
            'refresh_token': credentials.refresh_token,
            'token_uri': credentials.token_uri,
            'client_id': credentials.client_id,
            'client_secret': credentials.client_secret,
            'scopes': credentials.scopes}

if __name__ == "__main__":
    app.run(host='0.0.0.0', port=5000)