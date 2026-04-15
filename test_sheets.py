"""
Google Sheets API テストスクリプト
保存ログシートにテスト行を1行追加する
"""

import os
from datetime import datetime

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
CREDENTIALS_PATH = os.path.expanduser("~/law-secretary/secrets/oauth_credentials.json")
TOKEN_PATH = os.path.expanduser("~/law-secretary/secrets/token.json")
SPREADSHEET_ID = "1-dZd7iC2-eXLCUOwHGNZCDV8U7y9ECL6wNBk5YC-czc"
SHEET_NAME = "保存ログ"


def get_credentials():
    """OAuth2.0認証情報を取得する。"""
    creds = None
    if os.path.exists(TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_PATH, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_PATH, "w") as token:
            token.write(creds.to_json())
    return creds


def main():
    creds = get_credentials()
    service = build("sheets", "v4", credentials=creds)

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    row = [now, "テスト", "秘書エージェント", "test.pdf", "01_事件記録/テスト", False, False, False, False]

    body = {"values": [row]}
    result = (
        service.spreadsheets()
        .values()
        .append(
            spreadsheetId=SPREADSHEET_ID,
            range=f"{SHEET_NAME}!A:L",
            valueInputOption="USER_ENTERED",
            insertDataOption="INSERT_ROWS",
            body=body,
        )
        .execute()
    )

    print(f"追加完了: {result.get('updates', {}).get('updatedRange', '')}")


if __name__ == "__main__":
    main()
