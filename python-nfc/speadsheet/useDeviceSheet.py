import gspread
from google.oauth2 import service_account
from datetime import datetime
import socket

# サービスアカウントキーへのパス
SERVICE_ACCOUNT_FILE = R'C:\Users\oonuk\OneDrive\デスクトップ\python-nfc\speadsheet\speadsheet-384209-a8ef9c130ab1.json'

# スプレッドシートのID
SPREADSHEET_ID = '1rS-Dx1coBTUitNA8ZXioWy5qWtlnkv48gDmJHIxybHo'

# 認証
scopes = ["https://www.googleapis.com/auth/spreadsheets"]
creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
client = gspread.authorize(creds)

# スプレッドシートを開く
spreadsheet = client.open_by_key(SPREADSHEET_ID)

# デバイスIDを取得（ここではホスト名を使用）
device_id = socket.gethostname()

# シートが存在するか確認し、なければ新規作成
try:
    sheet = spreadsheet.worksheet(device_id)
except gspread.WorksheetNotFound:
    sheet = spreadsheet.add_worksheet(title=device_id, rows="100", cols="10")

def find_empty_cell(sheet, column=1):
    row = 1
    while sheet.cell(row, column).value:
        row += 1
    return row

# 現在の日付と時刻を取得
current_time = datetime.now()
date_str = current_time.strftime('%Y-%m-%d')
time_str = current_time.strftime('%H:%M:%S')

# 空のセルを見つけて出勤記録を書き込む
row_to_write = find_empty_cell(sheet)
sheet.update_cell(row_to_write, 1, date_str)
sheet.update_cell(row_to_write, 2, time_str)
sheet.update_cell(row_to_write, 3, device_id)