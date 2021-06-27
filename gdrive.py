import gspread
from oauth2client.service_account import ServiceAccountCredentials
import csv

json_file = 'service_account.json'
file_name = 'Python_Test'
sheet_name1 = 'シート1'
sheet_name2 = 'csv_sheet'
csv_file_name = 'Davis.csv'

scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive']

# スプレッドシートにアクセス
credentials = ServiceAccountCredentials.from_json_keyfile_name(json_file, scope)
gc = gspread.authorize(credentials)
# ファイルオープン
workbook = gc.open(file_name)

# シートの作成
sheet = workbook.add_worksheet(title = 'new_worksheet', rows = '100', cols = '30')

sheet_list = [ws.title for ws in workbook.worksheets()]
print(sheet_list)