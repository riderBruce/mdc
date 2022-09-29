import numpy as np
import win32com.client
import re
import os

from model_data import *

class ExcelDataConverter:
    def __init__(self, fileName):
        self.fileName = fileName

    def is_excel_file(self):
        filename_tmp, file_extension = os.path.splitext(self.fileName)
        if file_extension in [".xlsx", ".xls"]:
            return True
        else:
            logWrite('[▷ fail           ] : 엑셀파일이 아닙니다. ' + self.fileName)
            return False

    def read_excel_usedRangeData(self, sheetNum):
        excel = win32com.client.dynamic.Dispatch('Excel.Application')
        excel.Visible = False
        # excel.ScreenUpdating = False
        # excel.AskToUpdateLinks = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(self.fileName)
        # sheetName = "sheet1"
        # sheet = wb.Sheets(sheetName)
        try:
            sheet = wb.Sheets(sheetNum)
        except Exception as ex:
            logWrite(f'{sheetNum} : 해당 시트가 없습니다. {ex}')
            return False
        # get all data from excel used range
        sheetName = sheet.Name
        usedRangeData = sheet.UsedRange()
        # close excel without save
        wb.Close(False)
        excel.quit()

        return sheetName, usedRangeData

    def is_pension_file(self, usedRangeData):
        if usedRangeData[0][0] == "No." and usedRangeData[0][1] in ["근로년월", "신고상태"]:
            if usedRangeData[2][0] == 1:
                return True
            else:
                logWrite('[▷ fail           ] : 퇴직공제부금 파일이지만 데이터가 없습니다. ' + self.fileName)
                return False
        else:
            logWrite('[▷ fail           ] : 퇴직공제부금 파일이 아닙니다. ' + self.fileName)
            return False

    def convert_pension_data_for_DB(self, sheetName, send_date, usedRangeData):
        filename = self.fileName
        filename_stem = Path(filename).stem
        # get Table range
        usedRangeTable = usedRangeData
        # get Table head and body
        usedRangeTableHead = usedRangeTable[:2]
        usedRangeTableBody = usedRangeTable[2:]
        # columnNames
        columnNames = [""] * len(usedRangeTableHead[0])
        for row in usedRangeTableHead:
            for i, val in enumerate(row):
                if val is None:
                    val = ""
                columnNames[i] += str(val)

        df = pd.DataFrame(usedRangeTableBody, columns=columnNames)

        # 근로년월 데이터가 있는 경우
        if '근로년월' in columnNames:
            columns = ['No.', '근로년월', '공제가입번호', '계약명', '소속', '성명', '직종', '확정일수', '수정일']
            df = df[columns]
            df['직종'] = df['직종'].fillna(" ")
            df = df.groupby(['근로년월', '공제가입번호', '계약명', '소속', '직종', '수정일']).agg(
                {'확정일수': 'sum', '성명': 'count'}
            )
            df = df.reset_index()

            if df.empty:
                return df, '0000-00'

            df['근로년월'] = df['근로년월'].apply(lambda x: str(x)[:4]+'-'+str(x)[4:6])

        # 근로년월 데이터가 없는 경우 : 받은달 전달 기준
        else:
            columns = ['No.', '공제가입번호', '계약명', '소속', '성명', '직종', '확정일수', '수정일']
            df = df[columns]
            df['직종'] = df['직종'].fillna(" ")
            df = df.groupby(['공제가입번호', '계약명', '소속', '직종', '수정일']).agg(
                {'확정일수': 'sum', '성명': 'count'}
            )
            df = df.reset_index()

            if df.empty:
                return df, '0000-00'

            df['근로년월'] = datetime.strftime(datetime.strptime(send_date[:10], '%Y-%m-%d') - relativedelta(months=1),
                                           '%Y-%m')

        df['수정일'] = df['수정일'].apply(lambda x: str(x)[:8])
        df['계약명'] = df['계약명'].apply(lambda x: str(x).strip())
        df['sheet_name'] = sheetName
        df['attachment'] = filename_stem
        df['send_date'] = send_date

        target_month = max(df['근로년월'].unique())

        df = df.rename(columns={
            '근로년월': '근로년월',
            '계약명': '현장명p',
            '소속': '업체명p',
            '직종': '직종',
            '공제가입번호': '공제가입번호',
            '수정일': '수정일',
            '확정일수': '확정일수',
            '성명': '인원수',
            'sheet_name': 'sheet_name',
            'attachment': 'attachment',
            'send_date': 'send_date'
        })

        return df, target_month


if __name__ == "__main__":




