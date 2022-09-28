from model_data import *
import win32com.client as win32
import pandas as pd


# Get Excel Data
fileName = r"D:\Project_Python\webMDChecker\MDChecker\data\220902\전자카드 시스템 현장 등록 현황_020831.xlsx"
excel = win32.dynamic.Dispatch('Excel.Application')
excel.Visible = False
excel.DisplayAlerts = False

wb = excel.Workbooks.Open(fileName)
sheet = wb.Sheets(2)

usedRangeData = sheet.UsedRange()

head_row_number = 0
for num, row in enumerate(usedRangeData):
    for cell in row:
        if cell in ('공제가입번호', '현장코드', '비고'):
            head_row_number = num
            break
columnNames = usedRangeData[head_row_number]
data = usedRangeData[head_row_number+1:]

df = pd.DataFrame(data, columns=columnNames)

df = df[['현장코드','공제회 등록 현장명']]
df = df.rename(columns={
    '현장코드': '현장코드',
    '공제회 등록 현장명': '현장명p'
})

df = df.dropna(how="any")
df = df[df['현장코드'].str.len() == 4]
df = df.drop_duplicates(subset=['현장명p'])
df = df.drop_duplicates()

print(df)

# input DB
dc = DataControl('SERVER')

try:
    dc.insert_new_site_name_to_mdc_mst_site(df)
except Exception as ex:
    print(ex)






