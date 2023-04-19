import nfc
from nfc.clf import RemoteTarget
import datetime
from oauth2client.service_account import ServiceAccountCredentials
import gspread

pre_connect_time1 = datetime.datetime.now() - datetime.timedelta(days=1)
pre_connect_time2 = datetime.datetime.now() - datetime.timedelta(days=1)

def on_connect(tag):
    global pre_connect_time1,pre_connect_time2
    dt_now = datetime.datetime.now()
    dt1 = dt_now - pre_connect_time1
    dt2 = dt_now - pre_connect_time2
    if(tag.identifier==b'...自分のカードのIDｍにすればよいです'):
        if(dt1.total_seconds()>60*30):
            pre_connect_time1 = dt_now
            print("出勤")
            write_spread_sheet(["出勤",dt_now.strftime('%Y-%m-%d %H:%M:%S')])
    elif(tag.identifier==b'...自分のカードのIDｍにすればよいです'):
        if (dt2.total_seconds() > 60 * 30):
            pre_connect_time2 = dt_now
            print("退勤")
            write_spread_sheet(["退勤", dt_now.strftime('%Y-%m-%d %H:%M:%S')])
    else:
        print("カードが違います")

def write_spread_sheet(str_contents):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_name('GCPのキーファイル', scope)
    gc = gspread.authorize(credentials)
    worksheet = gc.open_by_url(
        "スプレッドシートのURL").sheet1
    worksheet.append_row(str_contents)

if __name__=="__main__":
    clf = nfc.ContactlessFrontend('usb')
    while clf.connect(rdwr={'on-connect': on_connect, 'targets': ['212F']}):
        print("connect")
    clf.close()