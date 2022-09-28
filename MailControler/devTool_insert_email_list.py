from model_data import *
import win32com.client as win32
import pandas as pd


# data
fileName = r"D:\Project_Python\webMDChecker\MDChecker\data\220922\복사본 전자카드 시스템 현장 등록 현황_020831_기준.xlsx"
sheet_number = 1

# Get Excel Data
excel = win32.dynamic.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

wb = excel.Workbooks.Open(fileName)
sheet = wb.Sheets(sheet_number)

usedRangeData = sheet.UsedRange()

head_row_number = 0
for num, row in enumerate(usedRangeData):
    for cell in row:
        if cell in ('이름', '현장코드', '메일주소'):
            head_row_number = num
            break
columnNames = usedRangeData[head_row_number]
data = usedRangeData[head_row_number+1:]

df = pd.DataFrame(data, columns=columnNames)

df = df[['이름','현장코드', '메일주소']]
# df = df.rename(columns={
#     '현장코드': '현장코드',
#     '공제회 등록 현장명': '현장명p'
# })

df = df.dropna(how="any")
df = df[df['현장코드'].str.len() == 4]
# df = df.drop_duplicates(subset=['메일주소'])
df = df.drop_duplicates()

print(df)

# input DB
dc = DataControl('SERVER')

try:
    success = dc.insert_data_to_db(df, 'mdc_address')
    if success:
        dc.conn.commit()
        print("업로드 완료 되었습니다.")
except Exception as ex:
    print(ex)
