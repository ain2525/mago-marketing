import os
import json
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

def main():
    # 1. GitHub Secretsから認証情報を取得
    # "GOOGLE_SHEETS_CREDENTIALS" はGitHubで設定した名前と完全一致させる
    creds_json = os.environ.get("GOOGLE_SHEETS_CREDENTIALS")
    
    if not creds_json:
        raise ValueError("Error: Secrets 'GOOGLE_SHEETS_CREDENTIALS' not found.")

    creds_dict = json.loads(creds_json)
    
    # 2. Googleスプレッドシートへの認証設定
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    client = gspread.authorize(creds)

    # 3. スプレッドシートを開く
    # 1dJwYYK-koOgU0V9z83hfz-Wjjjl_UNbl_N6eHQk5OmI
    # 例: https://docs.google.com/spreadsheets/d/abc1234567/edit
    spreadsheet_url = "1dJwYYK-koOgU0V9z83hfz-Wjjjl_UNbl_N6eHQk5OmI" 
    
    try:
        workbook = client.open_by_url(spreadsheet_url)
        sheet = workbook.sheet1  # 1枚目のシートを選択
        
        # 4. テスト書き込み（現在時刻とメッセージ）
        current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        data = [current_time, "GitHub Actions接続テスト成功！", "広瀬さん連携用"]
        
        sheet.append_row(data)
        print(f"Success: Added row {data}")

    except Exception as e:
        print(f"Error: {e}")
        raise e

if __name__ == "__main__":
    main()
