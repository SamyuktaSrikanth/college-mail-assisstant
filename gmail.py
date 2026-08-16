import os.path
import base64
from email.message import EmailMessage
# pyrefly: ignore [missing-import]
from bs4 import BeautifulSoup

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

from email.message import EmailMessage
import base64

from googleapiclient.discovery import build

# If modifying these scopes, delete the file token.json.
SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.send"
]

CREDENTIALS_FILE = 'credentials/credentials.json'
TOKEN_FILE = 'credentials/token.json'

def get_gmail_service():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())

    try:
        # Call the Gmail API
        service = build('gmail', 'v1', credentials=creds)
        return service
    except HttpError as error:
        # TODO(developer) - Handle errors from gmail API.
        print(f'An error occurred: {error}')
        return None

def fetch_unread_emails(service, max_results=10):
    """Fetches unread emails from the inbox."""
    try:
        results = service.users().messages().list(userId='me', labelIds=['INBOX', 'UNREAD'], maxResults=max_results).execute()
        messages = results.get('messages', [])

        if not messages:
            print('No new messages found.')
            return []

        parsed_emails = []
        for message in messages:
            msg = service.users().messages().get(userId='me', id=message['id'], format='full').execute()
            
            # Extract headers
            payload = msg['payload']
            headers = payload.get('headers', [])
            
            subject = ''
            sender = ''
            date = ''
            
            for d in headers:
                if d['name'] == 'Subject':
                    subject = d['value']
                if d['name'] == 'From':
                    sender = d['value']
                if d['name'] == 'Date':
                    date = d['value']

            # Extract body
            body = ""
            if 'parts' in payload:
                for part in payload['parts']:
                    if part['mimeType'] == 'text/plain':
                        data = part['body'].get('data')
                        if data:
                            body = base64.urlsafe_b64decode(data).decode('utf-8')
                    elif part['mimeType'] == 'text/html':
                        data = part['body'].get('data')
                        if data:
                            html_body = base64.urlsafe_b64decode(data).decode('utf-8')
                            soup = BeautifulSoup(html_body, 'html.parser')
                            body = soup.get_text(separator='\n') # Fallback to HTML if no plain text
            else:
                data = payload['body'].get('data')
                if data:
                    body = base64.urlsafe_b64decode(data).decode('utf-8')

            parsed_emails.append({
                'id': msg['id'],
                'subject': subject,
                'sender': sender,
                'date': date,
                'body': body,
                'link': f"https://mail.google.com/mail/u/0/#inbox/{msg['id']}"
            })
            
        return parsed_emails

    except HttpError as error:
        print(f'An error occurred: {error}')
        return []

if __name__ == '__main__':
    print("Authenticating with Gmail API...")
    service = get_gmail_service()
    if service:
        print("Authentication successful!")
        print("Fetching up to 5 unread emails to test...")
        emails = fetch_unread_emails(service, max_results=5)
        for i, email in enumerate(emails, 1):
            print(f"--- Email {i} ---")
            print(f"Subject: {email['subject']}")
            print(f"From: {email['sender']}")
            print(f"Date: {email['date']}")
            print(f"Link: {email['link']}")
            print(f"Body snippet: {email['body'][:100]}...\n")


def send_email(service, recipient, subject, html_body):

    message = EmailMessage()

    message["To"] = recipient
    message["Subject"] = subject

    message.set_content(
        "Please view this email in an HTML-compatible mail client."
    )

    message.add_alternative(
        html_body,
        subtype="html"
    )

    encoded_message = base64.urlsafe_b64encode(
        message.as_bytes()
    ).decode()

    body = {
        "raw": encoded_message
    }

    result = (
        service.users()
        .messages()
        .send(
            userId="me",
            body=body
        )
        .execute()
    )

    return result
