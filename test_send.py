from email.message import EmailMessage
import base64

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build


SCOPES = [
    "https://www.googleapis.com/auth/gmail.readonly",
    "https://www.googleapis.com/auth/gmail.send",
]


# Put YOUR Gmail address here
MY_EMAIL = "samyukta.srikanth2023@vitstudent.ac.in"


def get_gmail_service():

    creds = None

    try:
        creds = Credentials.from_authorized_user_file(
            "credentials/token.json",
            SCOPES
        )
    except Exception:
        pass

    if not creds or not creds.valid:

        if creds and creds.expired and creds.refresh_token:

            creds.refresh(Request())

        else:

            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials/credentials.json",
                SCOPES
            )

            creds = flow.run_local_server(
                port=0
            )

        with open(
            "credentials/token.json",
            "w"
        ) as token:

            token.write(
                creds.to_json()
            )

    return build(
        "gmail",
        "v1",
        credentials=creds
    )


def send_test_email(service):

    message = EmailMessage()

    message["To"] = MY_EMAIL
    message["Subject"] = (
        "College Mail Assistant - TEST"
    )

    message.set_content(
        """
Hello!

This is a test email from my
College Mail Assistant.

If you received this, Gmail API
sending is working successfully.

Next step: automated 3-hour summaries.

- College Mail Assistant
"""
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

    print(
        "Email sent successfully!"
    )

    print(
        "Message ID:",
        result["id"]
    )


def main():

    print("Connecting to Gmail...")

    service = get_gmail_service()

    print("Connected.")

    send_test_email(service)


if __name__ == "__main__":
    main()