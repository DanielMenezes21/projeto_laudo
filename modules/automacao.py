import os.path
import os
import webbrowser
import base64
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import email
from google.auth.transport.requests import Request
import re

SCOPES = ['https://mail.google.com/']

class EmailAutomator:
    def login_gmail():
        creds = None
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(r'modules\credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            with open('token.json', 'w') as token:
                token.write(creds.to_json())
        return build('gmail', 'v1', credentials=creds)

    def baixar_anexos(service, query):
        unread_query = f"{query} is:unread" if query else "is:unread"
        results = service.users().messages().list(userId='me', q=unread_query).execute()
        messages = results.get('messages', [])

        for msg in messages:
            msg_id = msg['id']
            message = service.users().messages().get(userId='me', id=msg_id).execute()
            payload = message.get('payload', {})

            headers = payload.get('headers', [])
            subject = "No_Subject"
            for header in headers:
                if header.get("name", "").lower() == "subject":
                    subject = header.get("value")
                    break

            safe_subject = re.sub(r'[\\/*?:"<>|]', "_", subject)
            folder_path = os.path.join("anexos", safe_subject)
            os.makedirs(folder_path, exist_ok=True)

            parts = payload.get('parts', [])
            for part in parts:
                filename = part.get("filename")
                if filename:
                    body = part.get("body", {})
                    att_id = body.get("attachmentId")
                    if att_id:
                        att = service.users().messages().attachments().get(
                            userId='me', messageId=msg_id, id=att_id
                        ).execute()
                        data = base64.urlsafe_b64decode(att['data'].encode('UTF-8'))
                        path = os.path.join(folder_path, filename)
                        with open(path, "wb") as f:
                            f.write(data)
                        print(f"✔️ Anexo salvo: {path}")

#550113588096-nilco8h5n1iue2uvsr35uukopdkv5fcd.apps.googleusercontent.com