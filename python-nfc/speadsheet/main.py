import gspread
import json
import subprocess
import socket

from google.oauth2.service_account import Credentials
from datetime import datetime


secret_credentials_json_oath =  R'C:\Users\oonuk\OneDrive\デスクトップ\python-nfc\speadsheet\speadsheet-384209-a8ef9c130ab1.json' 

scopes = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

credentials = Credentials.from_service_account_file(
    secret_credentials_json_oath,
    scopes=scopes
)

gc = gspread.authorize(credentials)
workbook = gc.open_by_key('1rS-Dx1coBTUitNA8ZXioWy5qWtlnkv48gDmJHIxybHo')
sheet = workbook.sheet1

# Update a range of cells using the top left corner address
# worksheet.update('A2', [[1, 2], [3, 4]])

# # Or update a single cell
# worksheet.update('B5', "it's down there somewhere, let me take another look.")

# # Format the header
# worksheet.format('A2:B2', {'textFormat': {'bold': True}})

# # A1からM8までの範囲を取得
# print(worksheet.get('A2'))


# def find_empty_cell(sheet, column=1):
#     row = 1
#     while sheet.cell(row, column).value:
#         row += 1
#     return row

# # 10回繰り返してデータを書き込む
# for _ in range(1):
#     row_to_write = find_empty_cell(sheet)
#     sheet.update_cell(row_to_write, 1, 'your_data')

def find_empty_cell(sheet, column=1):
    row = 1
    while sheet.cell(row, column).value:
        row += 1
    return row

def get_serial_number():
    try:
        serial = subprocess.check_output("cat /proc/cpuinfo | grep Serial | cut -d ' ' -f 2", shell=True).decode("utf-8").rstrip()
    except Exception as e:
        serial = "Error: {}".format(e)
    return serial


# 現在の日付と時刻を取得
current_time = datetime.now()
date_str = current_time.strftime('%Y-%m-%d')
time_str = current_time.strftime('%H:%M:%S')

# デバイスのIDで行う場合
device_id = socket.gethostname()
# Raspberry Piのシリアル番号を取得
# device_id = get_serial_number()

# 空のセルを見つけて出勤記録を書き込む
row_to_write = find_empty_cell(sheet)
sheet.update_cell(row_to_write, 1, date_str)
sheet.update_cell(row_to_write, 2, time_str)
sheet.update_cell(row_to_write, 3, device_id)

