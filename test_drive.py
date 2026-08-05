from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
import json
with open('token.json', 'r') as f:
    creds_data = json.load(f)
creds = Credentials.from_authorized_user_info(creds_data)
service = build('drive', 'v3', credentials=creds)
res = service.files().list(q="name = 'Carrier_Statements'", corpora='allDrives', includeItemsFromAllDrives=True, supportsAllDrives=True, fields='files(id, name)').execute()
print('Exact search:', res.get('files', []))
