import gspread
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.errors import HttpError
from email.message import EmailMessage
import base64
import pickle
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CREDS_PATH = os.path.join(BASE_DIR, "credentials.json")
TOKEN_PATH = os.path.join(BASE_DIR, "token.pkl")

def authenticate_google_services():
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive.readonly", "https://www.googleapis.com/auth/gmail.send"]
    creds = None

    if os.path.exists(TOKEN_PATH):
        with open(TOKEN_PATH, 'rb') as token:
            creds = pickle.load(token)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDS_PATH, scope)
            creds = flow.run_local_server(port=0)

        with open(TOKEN_PATH, 'wb') as token:
            pickle.dump(creds, token)

    client = gspread.authorize(creds)
    drive_service = build('drive', 'v3', credentials=creds)
    mail_service = build('gmail', 'v1', credentials=creds)

    return client, drive_service, mail_service

def create_google_sheet(sheet_name, snipeit_has_errors, share_with_email=None):
    client, _, _ = authenticate_google_services()
    spreadsheet = client.create(sheet_name)
    print(f"Generated {sheet_name}.xlsx.")

    if share_with_email and snipeit_has_errors:
        spreadsheet.share(share_with_email, perm_type='user', role='writer')
        print(f"Shared the sheet with {share_with_email}")

    return spreadsheet

def write_to_google_sheet(spreadsheet, sheet_name, df=None):
    worksheet = spreadsheet.add_worksheet(title=sheet_name, rows="100", cols="20")
    if df is not None:
        df_clean = df.fillna("")
        worksheet.update([df_clean.columns.values.tolist()] + df_clean.values.tolist())

    return worksheet

def download_google_sheet(spreadsheet, file_name):
    _, drive_service, _ = authenticate_google_services()
    file_id = spreadsheet.id
    file_path = os.path.join(BASE_DIR, f"{file_name}")

    request = drive_service.files().export_media(fileId=file_id, mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    with open(file_path, 'wb') as f:
        f.write(request.execute())
    print(f"Downloaded {file_name} to {file_path}")

    return file_path

def share_google_sheet(file_path, to_email):
    _, _, mail_service = authenticate_google_services()
    file_name = os.path.basename(file_path)
    try:
        message = EmailMessage()
        message.set_content("See attached file for final output")
        with open(file_path, 'rb') as content_file:
            content = content_file.read()
            message.add_attachment(content, maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=file_name)

        message['To'] = to_email
        message['Subject'] = os.path.splitext(file_name)[0]
        encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode()

        create_message = {
            'raw': encoded_message
        }
        mail_service.users().messages().send(userId="me", body=create_message).execute()
        print(f"Sent an email to {to_email} with {os.path.splitext(file_name)[0]} attached")
    except HttpError as error:
        print(f"An error occurred: {error}")

def delete_google_sheet(spreadsheet, sheet_name):
    _, service, _ = authenticate_google_services()

    service.files().delete(fileId=spreadsheet.id).execute()
    print(f"Deleted {sheet_name}")
