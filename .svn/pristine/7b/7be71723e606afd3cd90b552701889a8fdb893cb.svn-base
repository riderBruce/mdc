import win32com.client
from model_data import *

class ExcelFormMaker:
    def __init__(self, outputFileName, siteCode, bms_site_name, month):
        self.outputFileName = outputFileName
        self.siteCode = siteCode
        self.bms_site_name = bms_site_name
        self.f_date = datetime.strptime(month, '%Y-%m').strftime('%Y.%m.%d')
        self.t_date = (datetime.strptime(month, '%Y-%m') + relativedelta(months=1, days=-1)).strftime('%Y.%m.%d')
        excel = win32com.client.dynamic.Dispatch('Excel.Application')
        self.wb = excel.Workbooks.Open(self.outputFileName)
        excel.Visible = False
        excel.DisplayAlerts = False
        # 총괄표
        self.styling_excel_form_summary(1)
        # 누계표 - 금번 첫 현장발송시(9/23 예정)에는 누계표 제외하고 비교표만 보낼 예정
        # self.styling_excel_form_cummsum(2)
        logWrite("[엑셀 Formatter    ] : 적용 완료")
        self.wb.Save()
        self.wb.Close(False)

    def colnum_string(self, n):
        # A:65 ~ Z:90 - 26자
        string = ""
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            string = chr(65 + remainder) + string
        return string

    def styling_excel_form_summary(self, num):
        ws = self.wb.Sheets(num)

        ws.Columns(1).Delete()

        # Full Table
        ws.UsedRange.Font.Size = 10

        ws.UsedRange.Interior.ColorIndex = 2 # white
        ws.UsedRange.Borders.ColorIndex = 1
        ws.UsedRange.Borders.Weight = 2
        ws.UsedRange.Borders.LineStyle = 1
        ws.UsedRange.Font.Bold = False
        ws.UsedRange.RowHeight = 23
        ws.UsedRange.ColumnWidth = 13
        ws.UsedRange.VerticalAlignment = 2

        # 열너비
        ws.Columns(1).ColumnWidth = 30
        ws.Columns(2).ColumnWidth = 10
        ws.Columns(3).ColumnWidth = 25
        ws.Columns(4).ColumnWidth = 10
        ws.Columns(5).ColumnWidth = 10
        ws.Columns(6).ColumnWidth = 10
        ws.Columns(7).ColumnWidth = 10
        ws.Columns(8).ColumnWidth = 10
        ws.Columns(9).ColumnWidth = 30

        # 정렬
        ws.Columns(1).HorizontalAlignment = 3  # 2:left, 3:center 4:right
        ws.Columns(2).HorizontalAlignment = 3  # 2:left, 3:center 4:right
        ws.Columns(3).HorizontalAlignment = 3  # 2:left, 3:center 4:right
        ws.Columns(4).HorizontalAlignment = 4  # 2:left, 3:center 4:right
        ws.Columns(5).HorizontalAlignment = 4  # 2:left, 3:center 4:right
        ws.Columns(6).HorizontalAlignment = 4  # 2:left, 3:center 4:right
        ws.Columns(7).HorizontalAlignment = 4  # 2:left, 3:center 4:right
        ws.Columns(8).HorizontalAlignment = 3  # 2:left, 3:center 4:right
        ws.Columns(9).HorizontalAlignment = 2  # 2:left, 3:center 4:right
        ws.Rows(1).HorizontalAlignment = 3

        # 숫자열
        for i in range(4, 8):
            ws.Columns(i).NumberFormat = '#,#### '  # 숫자가 아닌 셀이 있어도 숫자셀만 변형됨



        # header row
        ws.Rows(2).Insert(1)
        ws.Range('A1:A2').Merge()
        ws.Range('B1').Value = '업체'
        ws.Range('B1:C1').Merge()
        ws.Range('B2').Value = '구분'
        ws.Range('C2').Value = '업체명'
        ws.Range('D1:F1').Merge()
        ws.Range('D1').Value = '당사 전자 작업일보\n(HPMS/HCM, Easy 작업일보)'
        ws.Range('D2').Value = '근로자수(a)\n(소장/직원\n미포함)'
        ws.Range('E2').Value = '※ 소장'
        ws.Range('F2').Value = '※ 직원'
        ws.Range('G1').Value = '전자카드\n(공제회)'
        ws.Range('G2').Value = '근로자수(b)\n(소장/직원\n미포함)'
        ws.Range('H1:H2').Merge()
        ws.Range('H1').Value = '대비\n(b/a)'
        ws.Range('I1:I2').Merge()

        # 헤더행 높이
        ws.Rows(1).RowHeight = 40
        ws.Rows(2).RowHeight = 40

        # 헤더행 글자 크기
        ws.Range('A1:I2').Font.Bold = True
        ws.Range('B2:G2').Font.Size = 8

        # 본문 글자 크기

        # col/row number
        nCol = ws.UsedRange.Columns.Count  # column수 확인
        nRow = ws.UsedRange.Rows.Count  # row수 확인

        ws.Range(f'A3:A{nRow-1}').Merge()
        ws.Range('A3').WrapText = True
        ws.Range(f'B4:B{nRow-1}').Merge()

        ws.Range(f'A{nRow}:C{nRow}').Merge()
        ws.Range(f'A{nRow}').Value = '총 합계 (당사 + 협력업체)'

        headerAdd = f'A1:{self.colnum_string(nCol)}2'
        ws.Range(headerAdd).Interior.ColorIndex = 19

        footerAdd = f'A{nRow}:{self.colnum_string(nCol)}{nRow}'
        ws.Range(footerAdd).Interior.ColorIndex = 19
        ws.Range(footerAdd).Font.Bold = True

        ws.Range(f'D2:D{nRow-1}').Interior.ColorIndex = 35
        ws.Range(f'D2:D{nRow-1}').Font.Bold = True
        ws.Range(f'G2:G{nRow-1}').Interior.ColorIndex = 35
        ws.Range(f'G2:G{nRow-1}').Font.Bold = True

        # 소장 / 직원 숫자 괄호 넣기
        ws.Range(f'E2:F{nRow-1}').NumberFormat = '(#,###) ;- (#,###) ; '
        ws.Range(f'E2:F{nRow-1}').Font.Size = 9

        # 대비 % 넣기
        ws.Range(f'H2:H{nRow}').NumberFormat = '#,###%; -#,###%; '
        ws.Range(f'H2:H{nRow}').Font.Size = 8

        # text
        ws.Range(f'A{nRow + 2}').Value \
            = '※ [용도]  퇴직공제부금 관련, 협력업체별 근로내역 확정전 원도급 관리자의 확인시 참고용'
        ws.Range(f'A{nRow + 2}').HorizontalAlignment = 2
        ws.Range(f'A{nRow + 2}').Font.Size = 9

        ws.Range(f'A{nRow + 3}').Value \
            = '※ [작업일보 근로자 수 (a)]  현장에서 해당월에 HPMS/HCM, Easy 작업일보에 입력한 근로자 수 기준'
        ws.Range(f'A{nRow + 3}').HorizontalAlignment = 2
        ws.Range(f'A{nRow + 3}').Font.Size = 9

        ws.Range(f'A{nRow + 4}').Value \
            = "※ [전자카드 근로자 수 (b)]  전자카드 근무관리시스템의 해당월 '근로내역' 기준"
        ws.Range(f'A{nRow + 4}').HorizontalAlignment = 2
        ws.Range(f'A{nRow + 4}').Font.Size = 9

        ws.Range(f'A{nRow + 5}').Value \
            = '※ [문의]  제도 관련 : 안전사업지원실 사업관리팀 (02-746-1940)  /  데이터 관련 : 예산관리실 RM팀 (02-746-3643, 1339, 2262, 2782)'
        ws.Range(f'A{nRow + 5}').HorizontalAlignment = 2
        ws.Range(f'A{nRow + 5}').Font.Size = 9

        # Top line 삽입
        ws.Rows(1).Insert()
        ws.Rows(1).Insert()
        ws.Rows(1).RowHeight = 60
        ws.Rows(1).VerticalAlignment = 2 # 가운데 정렬
        ws.Rows(1).HorizontalAlignment = 2
        ws.Range('A1').Value = '▶ 현장 출역정보 비교표 (당사 전자 작업일보 vs. 전자카드 근무관리시스템)'
        ws.Range('A1').Font.Bold = True
        ws.Range('A1').Font.Size = 16

        ws.Range('A2').Value = f"         {datetime.now().strftime('%y.%m.%d. %H:%M 기준')}"
        ws.Range('A2').Font.Color = "&hFF0000"
        ws.Range('A2').Font.Size = 9
        ws.Range('A2').HorizontalAlignment = 2

        ws.Range('I2').Value = f'[인원수 합산기간] {self.f_date}~{self.t_date} (인원단위 : 명)'
        ws.Range('I2').Font.Size = 9
        ws.Range('I2').HorizontalAlignment = 4

        # common number
        nCol = ws.UsedRange.Columns.Count  # column수 확인
        nRow = ws.UsedRange.Rows.Count  # row수 확인
        print_area = f'A1:{self.colnum_string(nCol)}{nRow}'
        ws.PageSetup.Zoom = False
        ws.PageSetup.FitToPagesTall = False
        ws.PageSetup.FitToPagesWide = 1
        ws.PageSetup.PrintArea = print_area
        ws.PageSetup.LeftMargin = 25
        ws.PageSetup.RightMargin = 25
        ws.PageSetup.TopMargin = 50
        ws.PageSetup.BottomMargin = 50
        ws.PageSetup.Orientation = 2 #가로로 출력 # 1: 세로 / 2:가로

    def styling_excel_form_cummsum(self, num):
        ws = self.wb.Sheets(num)

        ws.Rows(1).Copy()
        ws.Rows(1).Insert(1)

        # col/row number
        nCol = ws.UsedRange.Columns.Count
        nRow = ws.UsedRange.Rows.Count

        # Full Table
        ws.UsedRange.Font.Size = 10

        ws.UsedRange.Interior.ColorIndex = 2 # white
        ws.UsedRange.Borders.ColorIndex = 1
        ws.UsedRange.Borders.Weight = 2
        ws.UsedRange.Borders.LineStyle = 1
        ws.UsedRange.Font.Bold = False
        ws.UsedRange.RowHeight = 23
        ws.UsedRange.ColumnWidth = 13
        ws.UsedRange.VerticalAlignment = 2

        # 열너비
        ws.Columns(1).ColumnWidth = 10
        ws.Columns(2).ColumnWidth = 25
        ws.Columns(3).ColumnWidth = 25
        ws.Columns(4).ColumnWidth = 10
        ws.Columns(5).ColumnWidth = 10
        ws.Columns(6).ColumnWidth = 10
        ws.Columns(7).ColumnWidth = 10
        ws.Columns(8).ColumnWidth = 20

        # 정렬
        ws.Columns(1).HorizontalAlignment = 3  # 2:left, 3:center 4:right
        ws.Columns(2).HorizontalAlignment = 3  # 2:left, 3:center 4:right
        ws.Columns(3).HorizontalAlignment = 2  # 2:left, 3:center 4:right
        ws.Columns(4).HorizontalAlignment = 4  # 2:left, 3:center 4:right
        ws.Columns(5).HorizontalAlignment = 4  # 2:left, 3:center 4:right
        ws.Columns(6).HorizontalAlignment = 4  # 2:left, 3:center 4:right
        ws.Columns(7).HorizontalAlignment = 4  # 2:left, 3:center 4:right
        ws.Columns(8).HorizontalAlignment = 3  # 2:left, 3:center 4:right
        ws.Rows(1).HorizontalAlignment = 3
        ws.Rows(2).HorizontalAlignment = 3

        # 숫자열
        for i in [4, 5, 6, 7]:
            ws.Columns(i).NumberFormat = '#,#### '  # 숫자가 아닌 셀이 있어도 숫자셀만 변형됨

        # data row
        ws.Range(f"G3:G{nRow}").Interior.ColorIndex = 35

        # header row
        headerAdd = f'A1:{self.colnum_string(nCol)}2'
        ws.Range(headerAdd).Interior.ColorIndex = 19
        ws.Range(headerAdd).Font.Bold = True
        ws.Rows(1).RowHeight = 40
        ws.Rows(2).RowHeight = 40

        ws.Range('A1:B1').Merge()
        ws.Range('A1').Value = '업체'

        ws.Range('C1:C2').Merge()
        ws.Range('C1').Value = '데이터 출처'

        ws.Range('D1:G1').Merge()
        ws.Range('D1').Value = '근로자 수 (※ 소장/직원 미포함)'

        ws.Range('H1:H2').Merge()


        # footer row
        footerAdd = f'A{nRow-3}:{self.colnum_string(nCol)}{nRow}'
        ws.Range(footerAdd).Interior.ColorIndex = 19
        ws.Range(footerAdd).Font.Bold = True

        ws.Range(f'A{nRow-3}:B{nRow-2}').Merge()
        ws.Range(f'A{nRow-1}:B{nRow}').Merge()

        # data row
        for i in range(1, nRow+1):
            if ws.Range(f"C{i}").Value == '퇴직공제부금':
                ws.Range(f"C{i}:G{i}").Interior.ColorIndex = 15

        # Top line 삽입
        ws.Rows(1).Insert()
        ws.Rows(1).Insert()
        ws.Rows(1).RowHeight = 60
        ws.Rows(1).VerticalAlignment = 2 # 가운데 정렬
        ws.Rows(1).HorizontalAlignment = 2
        ws.Range('A1').Value = f'▶ 현장 출역정보 누계표 (최근4개월) : [{self.siteCode}] {self.bms_site_name}'
        ws.Range('A1').Font.Bold = True
        ws.Range('A1').Font.Size = 16

        ws.Range('A2').Value = f"         {datetime.now().strftime('%y.%m.%d. %H:%M 기준')}"
        ws.Range('A2').Font.Color = "&hFF0000"
        ws.Range('A2').Font.Size = 9
        ws.Range('A2').HorizontalAlignment = 2


        # print
        nCol = ws.UsedRange.Columns.Count
        nRow = ws.UsedRange.Rows.Count
        print_area = f'A1:{self.colnum_string(nCol)}{nRow}'
        ws.PageSetup.Zoom = False
        ws.PageSetup.FitToPagesTall = False
        ws.PageSetup.FitToPagesWide = 1
        ws.PageSetup.PrintArea = print_area
        ws.PageSetup.LeftMargin = 25
        ws.PageSetup.RightMargin = 25
        ws.PageSetup.TopMargin = 50
        ws.PageSetup.BottomMargin = 50
        ws.PageSetup.Orientation = 1 #가로/세로 출력 # 1: 세로 / 2:가로



if __name__ == '__main__':
    # from model_excel_export import SummaryExcelData
    # dc = DataControl()
    # summaryExcelData = SummaryExcelData(dc)
    # outputFileName = summaryExcelData.name_output_excel_file()
    # summaryExcelData.consist_summary_excel_data(outputFileName)
    # outputFileName = r"D:\Project_Data\webMDChecker\Result_\Reply(GJ57_2022-06)_2022-07-28_1032 - 복사본.xlsx"
    # ExcelFormMaker(outputFileName, 'GJ57', '2022-06')
    # import os
    # os.startfile(outputFileName)

    dc = DataControl('SERVER')
    from model_excel_export import SummaryExcelData
    se = SummaryExcelData(dc, 'I73C', '2022-06') # 당사만 있음
    outputFileName = se.name_output_excel_file()
    se.excel_writer(outputFileName)
    ExcelFormMaker(outputFileName, 'I73C', '2022-06')
    os.startfile(outputFileName)

    se = SummaryExcelData(dc, 'GJ26', '2022-07')  # 둘다 있음
    outputFileName = se.name_output_excel_file()
    se.excel_writer(outputFileName)
    ExcelFormMaker(outputFileName, 'GJ26', '2022-07')
    os.startfile(outputFileName)

    se = SummaryExcelData(dc, 'GJ57', '2022-06') # 협력업체만 있음
    outputFileName = se.name_output_excel_file()
    se.excel_writer(outputFileName)
    ExcelFormMaker(outputFileName, 'GJ57', '2022-06')
    os.startfile(outputFileName)

    # import win32com.client
    # excel = win32com.client.dynamic.Dispatch('Excel.Application')
    # outputFileName = r"D:\Project_Data\JobReport\Result_\DailyReport()_2022-03-03_1333.xlsx"
    # wb = excel.Workbooks.Open(outputFileName)
    # excel.Visible = True
    # excel.DisplayAlerts = False
    # ws = wb.Sheets(1)

# ws.Rows(1).Insert(1)
# ws.Rows(1).RowHeight = 22
# ws.Rows(1).VerticalAlignment = 2  # 행 전체 1:top 2:center 3:bottom
# ws.Range('A1:A2').Merge()
# ws.Range('B1:B2').Merge()
# ws.Range('C1:C2').Merge()
# ws.Range('D1:H1').Merge()
# ws.Range('I1:M1').Merge()
#
# headerAdd = f'A1:{self.colnum_string(nCol)}2'
# ws.Range(headerAdd).Interior.ColorIndex = 15
# ws.Range(headerAdd).Borders.ColorIndex = 1
# ws.Range(headerAdd).Borders.Weight = 2
# ws.Range(headerAdd).Borders.LineStyle = 1
# ws.Range(headerAdd).HorizontalAlignment = 3
# ws.Range(headerAdd).Font.Bold = True
# ws.Range('D2:M2').Font.Size = 10
#
# ws.Rows(2).RowHeight = 30
#
# # subTotal
# for i in range(1, nRow + 1):
#     if ws.Range(f'B{i}').Value == 'Project Total':
#         ws.Range(f'B{i}:{self.colnum_string(idxNo)}{i}').Merge()
#         ws.Range(f'B{i}:{self.colnum_string(nCol)}{i}').Interior.ColorIndex = 35  # 연녹색
#
# # footer row
# ws.Range(f'A{nRow + 1}:{self.colnum_string(idxNo)}{nRow + 1}').Merge()
# footerAdd = f'A{nRow + 1}:{self.colnum_string(nCol)}{nRow + 1}'
# ws.Range(footerAdd).Interior.ColorIndex = 36  # 연노란색
#
# # 하단 합계행 복사하여 상단 헤더 아래에 넣기
# ws.Rows(nRow + 1).Copy()
# ws.Rows(3).Insert(1)
#
# # 제목 행에 값 넣기
# ws.Range('A1').Value = 'Project\nCode'
# ws.Range('B1').Value = 'Discipline'
# ws.Range('C1').Value = 'Subcon\nName'
# ws.Range('D1').Value = 'Man Power'
# ws.Range('I1').Value = 'Equipment'
# ws.Range('D2').Value = 'D-3\n(' + D3_Day + ')'
# ws.Range('E2').Value = 'D-2\n(' + D2_Day + ')'
# ws.Range('F2').Value = 'D-1\n(' + D1_Day + ')'
# ws.Range('G2').Value = 'Today\n(' + T_Date + ')'
# ws.Range('H2').Value = 'Accumulate'
# ws.Range('I2').Value = 'D-3\n(' + D3_Day + ')'
# ws.Range('J2').Value = 'D-2\n(' + D2_Day + ')'
# ws.Range('K2').Value = 'D-1\n(' + D1_Day + ')'
# ws.Range('L2').Value = 'Today\n(' + T_Date + ')'
# ws.Range('M2').Value = 'Accumulate'
#
# # 제목행 삽입
# ws.Rows(1).Insert(1)  # 행 추가
# ws.Rows(1).RowHeight = 40
# ws.Rows(1).VerticalAlignment = 2  # 행 전체 1:top 2:center 3:bottom
# ws.Range('A1').Font.Size = 18
# ws.Range('A1').Value = f"□ Summary Daily Report ({sRunTime}) "
# ws.Range('A1').Font.Bold = True

# elif sheetName == 'ManSum' or sheetName == 'EquipSum':
#     # 인덱스열 갯수
#     idxNo = 4
#     # 인덱스열
#     for i in range(1, idxNo+1):
#         ws.Columns(i).HorizontalAlignment = 2  # 왼쪽 정렬 : // 2:left, 3:center 4:right
#         ws.Columns(i).ColumnWidth = 20  # 모든 컬럼 넓이
#     ws.Columns(idxNo).ColumnWidth = 30
#     # 숫자열
#     for i in range(idxNo+1, nCol + 1):
#         ws.Columns(i).NumberFormat = '#,#### '  # 숫자가 아닌 셀이 있어도 숫자셀만 변형됨
#     # header row
#     headerAdd = f'A1:{self.colnum_string(nCol)}1'
#     ws.Range(headerAdd).Interior.ColorIndex = 15
#     ws.Range(headerAdd).Borders.ColorIndex = 1
#     ws.Range(headerAdd).Borders.Weight = 2
#     ws.Range(headerAdd).Borders.LineStyle = 1
#     ws.Range(headerAdd).HorizontalAlignment = 3
#     ws.Range(headerAdd).Font.Bold = True
#     ws.Rows(1).RowHeight = 60
#
#     # subTotal
#     for i in range(1, nRow+1):
#         if ws.Range(f'B{i}').Value == 'Project Total':
#             ws.Range(f'B{i}:{self.colnum_string(idxNo)}{i}').Merge()
#             ws.Range(f'B{i}:{self.colnum_string(nCol)}{i}').Interior.ColorIndex = 35 # 연녹색
#
#     # footer row
#     ws.Range(f'A{nRow}:{self.colnum_string(idxNo)}{nRow}').Merge()
#     footerAdd = f'A{nRow}:{self.colnum_string(nCol)}{nRow}'
#     ws.Range(footerAdd).Interior.ColorIndex = 36 # 노란색
#
#     # 하단 합계행 복사하여 상단 헤더 아래에 넣기
#     ws.Rows(nRow).Copy()
#     ws.Rows(2).Insert(1)
#
#     # 제목 행에 값 넣기
#     ws.Range('A1').Value = 'Project\nCode'
#     ws.Range('B1').Value = 'Discipline'
#     ws.Range('C1').Value = 'Subcon\nName'
#     ws.Range('D1').Value = 'ManpowerType'
#
#     ws.Range('E1').Value = 'D-3\n(' + D3_Day + ')'
#     ws.Range('F1').Value = 'D-2\n(' + D2_Day + ')'
#     ws.Range('G1').Value = 'D-1\n(' + D1_Day + ')'
#     ws.Range('H1').Value = 'Today\n(' + T_Date + ')'
#     ws.Range('I1').Value = 'Accumulate'
#
#     # 제목행
#     ws.Rows(1).Insert(1)  # 행 추가
#     ws.Rows(1).RowHeight = 40
#     ws.Rows(1).VerticalAlignment = 2  # 행 전체 1:top 2:center 3:bottom
#     ws.Range('A1').Font.Size = 18
#     ws.Range('A1').Font.Bold = True
#
#     if sheetName == 'ManSum':
#         ws.Range('A1').Value = f"□ Summary Manpower ({sRunTime}) "
#     elif sheetName == 'EquipSum':
#         ws.Range('A1').Value = f"□ Summary Equipment ({sRunTime}) "
#
# elif sheetName == 'Man':
#     # 인덱스열 갯수
#     idxNo = 10
#     # 인덱스열
#     for i in range(1, idxNo+1):
#         ws.Columns(i).HorizontalAlignment = 2  # 왼쪽 정렬 : // 2:left, 3:center 4:right
#         ws.Columns(i).ColumnWidth = 20  # 모든 컬럼 넓이
#     ws.Columns(idxNo).ColumnWidth = 30
#     # 숫자열
#     for i in range(idxNo+1, nCol + 1):
#         ws.Columns(i).NumberFormat = '#,#### '  # 숫자가 아닌 셀이 있어도 숫자셀만 변형됨
#     # header row
#     headerAdd = f'A1:{self.colnum_string(nCol)}1'
#     ws.Range(headerAdd).Interior.ColorIndex = 15
#     ws.Range(headerAdd).Borders.ColorIndex = 1
#     ws.Range(headerAdd).Borders.Weight = 2
#     ws.Range(headerAdd).Borders.LineStyle = 1
#     ws.Range(headerAdd).HorizontalAlignment = 3
#     ws.Range(headerAdd).Font.Bold = True
#     ws.Rows(1).RowHeight = 60
#
#     # subTotal
#     for i in range(1, nRow+1):
#         if ws.Range(f'B{i}').Value == 'Project Total':
#             ws.Range(f'B{i}:{self.colnum_string(idxNo)}{i}').Merge()
#             ws.Range(f'B{i}:{self.colnum_string(nCol)}{i}').Interior.ColorIndex = 35 # 연녹색
#     # footer row
#     ws.Range(f'A{nRow}:{self.colnum_string(idxNo)}{nRow}').Merge()
#     footerAdd = f'A{nRow}:{self.colnum_string(nCol)}{nRow}'
#     ws.Range(footerAdd).Interior.ColorIndex = 36 # 노란색
#
#     # 하단 합계행 복사하여 상단 헤더 아래에 넣기
#     ws.Rows(nRow).Copy()
#     ws.Rows(2).Insert(1)
#
#     # 제목 행에 값 넣기
#
#     # 제목행
#     ws.Rows(1).Insert(1)  # 행 추가
#     ws.Rows(1).RowHeight = 40
#     ws.Rows(1).VerticalAlignment = 2  # 행 전체 1:top 2:center 3:bottom
#     ws.Range('A1').Font.Size = 18
#     ws.Range('A1').Font.Bold = True
#
#     ws.Range('A1').Value = f"□ Daily Manpower ({sRunTime}) "
#
# elif sheetName == 'Equip':
#     # 인덱스열 갯수
#     idxNo = 9
#     # 인덱스열
#     for i in range(1, idxNo+1):
#         ws.Columns(i).HorizontalAlignment = 2  # 왼쪽 정렬 : // 2:left, 3:center 4:right
#         ws.Columns(i).ColumnWidth = 20  # 모든 컬럼 넓이
#     ws.Columns(idxNo).ColumnWidth = 30
#     # 숫자열
#     for i in range(idxNo+1, nCol + 1):
#         ws.Columns(i).NumberFormat = '#,#### '  # 숫자가 아닌 셀이 있어도 숫자셀만 변형됨
#     # header row
#     headerAdd = f'A1:{self.colnum_string(nCol)}1'
#     ws.Range(headerAdd).Interior.ColorIndex = 15
#     ws.Range(headerAdd).Borders.ColorIndex = 1
#     ws.Range(headerAdd).Borders.Weight = 2
#     ws.Range(headerAdd).Borders.LineStyle = 1
#     ws.Range(headerAdd).HorizontalAlignment = 3
#     ws.Range(headerAdd).Font.Bold = True
#     ws.Rows(1).RowHeight = 60
#     # subTotal
#     for i in range(1, nRow+1):
#         if ws.Range(f'B{i}').Value == 'Project Total':
#             ws.Range(f'B{i}:{self.colnum_string(idxNo)}{i}').Merge()
#             ws.Range(f'B{i}:{self.colnum_string(nCol)}{i}').Interior.ColorIndex = 35 # 연녹색
#     # footer row
#     ws.Range(f'A{nRow}:{self.colnum_string(idxNo)}{nRow}').Merge()
#     footerAdd = f'A{nRow}:{self.colnum_string(nCol)}{nRow}'
#     ws.Range(footerAdd).Interior.ColorIndex = 36 # 노란색
#
#     # 하단 합계행 복사하여 상단 헤더 아래에 넣기
#     ws.Rows(nRow).Copy()
#     ws.Rows(2).Insert(1)
#
#     # 제목 행에 값 넣기
#
#     # 제목행
#     ws.Rows(1).Insert(1)  # 행 추가
#     ws.Rows(1).RowHeight = 40
#     ws.Rows(1).VerticalAlignment = 2  # 행 전체 1:top 2:center 3:bottom
#     ws.Range('A1').Font.Size = 18
#     ws.Range('A1').Font.Bold = True
#
#     ws.Range('A1').Value = f"□ Daily Equipment ({sRunTime}) "
#
# elif sheetName == 'Qnt':
#     # 인덱스열 갯수
#     idxNo = 6
#     # 인덱스열
#     for i in range(1, idxNo+1):
#         ws.Columns(i).HorizontalAlignment = 2  # 왼쪽 정렬 : // 2:left, 3:center 4:right
#         ws.Columns(i).ColumnWidth = 20  # 모든 컬럼 넓이
#     ws.Columns(idxNo).ColumnWidth = 30
#     # 숫자열
#     for i in range(idxNo+1, nCol + 1):
#         ws.Columns(i).NumberFormat = '#,#### '  # 숫자가 아닌 셀이 있어도 숫자셀만 변형됨
#     # header row
#     headerAdd = f'A1:{self.colnum_string(nCol)}1'
#     ws.Range(headerAdd).Interior.ColorIndex = 15
#     ws.Range(headerAdd).Borders.ColorIndex = 1
#     ws.Range(headerAdd).Borders.Weight = 2
#     ws.Range(headerAdd).Borders.LineStyle = 1
#     ws.Range(headerAdd).HorizontalAlignment = 3
#     ws.Range(headerAdd).Font.Bold = True
#     ws.Rows(1).RowHeight = 60
#     # subTotal
#     for i in range(1, nRow+1):
#         if ws.Range(f'B{i}').Value == 'Project Total':
#             ws.Range(f'B{i}:{self.colnum_string(idxNo)}{i}').Merge()
#             ws.Range(f'B{i}:{self.colnum_string(nCol)}{i}').Interior.ColorIndex = 35 # 연녹색
#     # footer row
#     ws.Range(f'A{nRow}:{self.colnum_string(idxNo)}{nRow}').Merge()
#     footerAdd = f'A{nRow}:{self.colnum_string(nCol)}{nRow}'
#     ws.Range(footerAdd).Interior.ColorIndex = 36 # 노란색
#
#     # 하단 합계행 복사하여 상단 헤더 아래에 넣기
#     ws.Rows(nRow).Copy()
#     ws.Rows(2).Insert(1)
#
#     # 제목 행에 값 넣기
#
#     # 제목행
#     ws.Rows(1).Insert(1)  # 행 추가
#     ws.Rows(1).RowHeight = 40
#     ws.Rows(1).VerticalAlignment = 2  # 행 전체 1:top 2:center 3:bottom
#     ws.Range('A1').Font.Size = 18
#     ws.Range('A1').Font.Bold = True
#
#     ws.Range('A1').Value = f"□ Daily Quantity ({sRunTime}) "
#
# if sheetName in ['Prd', 'Prd_cum']:
#     # 인덱스열 갯수
#     idxNo = 7
#     # 인덱스열
#     for i in range(1, idxNo+1):
#         ws.Columns(i).HorizontalAlignment = 2  # 왼쪽 정렬 : // 2:left, 3:center 4:right
#         ws.Columns(i).ColumnWidth = 20  # 모든 컬럼 넓이
#     ws.Columns(idxNo).ColumnWidth = 30
#     # 숫자열
#     for i in range(idxNo+1, nCol + 1):
#         if ws.Range(f'{self.colnum_string(i)}2').Value in ['MAN', 'MAN_cum']:
#             ws.Columns(i).NumberFormat = '#,#### '  # 숫자가 아닌 셀이 있어도 숫자셀만 변형됨
#         elif ws.Range(f'{self.colnum_string(i)}2').Value in ['PRD', 'PRD_cum']:
#             ws.Columns(i).NumberFormat = '#,##0.00 ;-#,##0.00;'  # 숫자가 아닌 셀이 있어도 숫자셀만 변형됨
#             ws.Columns(i).Font.Bold = True
#         else:
#             ws.Columns(i).NumberFormat = '#,##0.00 ;-#,##0.00;'  # 숫자가 아닌 셀이 있어도 숫자셀만 변형됨
#
# # header row
#     ws.Range('A1:A3').Merge()
#     ws.Range('B1:B3').Merge()
#     ws.Range('C1:C3').Merge()
#     ws.Range('D1:D3').Merge()
#     ws.Range('E1:E3').Merge()
#     ws.Range('F1:F3').Merge()
#     ws.Range('G1:G3').Merge()
#
#     headerAdd = f'A1:{self.colnum_string(nCol)}2'
#     ws.Range(headerAdd).Interior.ColorIndex = 15
#     ws.Range(headerAdd).Borders.ColorIndex = 1
#     ws.Range(headerAdd).Borders.Weight = 2
#     ws.Range(headerAdd).Borders.LineStyle = 1
#     ws.Range(headerAdd).HorizontalAlignment = 3
#     ws.Range(headerAdd).Font.Bold = True
#
#     # subTotal
#     for i in range(1, nRow+1):
#         if ws.Range(f'B{i}').Value == 'Project Total':
#             ws.Range(f'B{i}:{self.colnum_string(idxNo)}{i}').Merge()
#             ws.Range(f'B{i}:{self.colnum_string(nCol)}{i}').Interior.ColorIndex = 35 # 연녹색
#
#     # footer row
#     ws.Range(f'A{nRow}:{self.colnum_string(idxNo)}{nRow}').Merge()
#     footerAdd = f'A{nRow}:{self.colnum_string(nCol)}{nRow}'
#     ws.Range(footerAdd).Interior.ColorIndex = 36 # 연노란색
#
#     # 하단 합계행 복사하여 상단 헤더 아래에 넣기
#     ws.Rows(nRow).Copy()
#     ws.Rows(4).Insert(1)
#     ws.Rows(3).EntireRow.Delete()
#
#     # 제목 행에 값 넣기
#     ws.Range('A1').Value = 'Project Code'
#     ws.Range('B1').Value = 'Discipline'
#     ws.Range('C1').Value = 'Subcontractor'
#     ws.Range('D1').Value = 'Category'
#     ws.Range('E1').Value = 'CWA'
#     ws.Range('F1').Value = 'CWP'
#     ws.Range('G1').Value = 'IWP'
#
#     # 제목행 삽입
#     ws.Rows(1).Insert(1)  # 행 추가
#     ws.Rows(1).RowHeight = 40
#     ws.Rows(1).VerticalAlignment = 2  # 행 전체 1:top 2:center 3:bottom
#     ws.Range('A1').Font.Size = 18
#     ws.Range('A1').Value = f"□ Summary Daily Productivity (by Direct Man Power, {sRunTime}) "
#     ws.Range('A1').Font.Bold = True
#
#     if sheetName == 'Prd':
#         ws.Range('A1').Value = f"□ Summary Daily Productivity (by Direct Man Power, {sRunTime}) "
#     elif sheetName == 'Prd_cum':
#         ws.Range('A1').Value = f"□ Summary Daily Accumulated Productivity (by Direct Man Power, {sRunTime}) "
#
# elif sheetName == 'Raw':
#     # 인덱스열 갯수
#     idxNo = 14
#     # 인덱스열
#     for i in range(1, idxNo+1):
#         ws.Columns(i).HorizontalAlignment = 2  # 왼쪽 정렬 : // 2:left, 3:center 4:right
#         ws.Columns(i).ColumnWidth = 10  # 모든 컬럼 넓이
#     ws.Columns(idxNo).ColumnWidth = 15
#     # 숫자열
#     for i in range(idxNo+1, nCol + 1):
#         ws.Columns(i).NumberFormat = '#,#### '  # 숫자가 아닌 셀이 있어도 숫자셀만 변형됨
#     # header row
#     headerAdd = f'A1:{self.colnum_string(nCol)}1'
#     ws.Range(headerAdd).Interior.ColorIndex = 15 # 회색
#     ws.Range(headerAdd).Borders.ColorIndex = 1
#     ws.Range(headerAdd).Borders.Weight = 2
#     ws.Range(headerAdd).Borders.LineStyle = 1
#     ws.Range(headerAdd).HorizontalAlignment = 3
#     ws.Range(headerAdd).Font.Bold = True
#     ws.Rows(1).RowHeight = 30
#     # footer row
#
#     # 하단 합계행 복사하여 상단 헤더 아래에 넣기
#
#     # 제목 행에 값 넣기
#
#     # 제목행
#     ws.Rows(1).Insert(1)  # 행 추가
#     ws.Rows(1).RowHeight = 40
#     ws.Rows(1).VerticalAlignment = 2  # 행 전체 1:top 2:center 3:bottom
#     ws.Range('A1').Font.Size = 18
#     ws.Range('A1').Font.Bold = True
#
#     ws.Range('A1').Value = f"□ Daily Raw Data ({sRunTime}) "


# 참고
# ws.Columns(1).HorizontalAlignment = 2  # 좌측정렬 : 인덱스 컬럼 A ~ D, 2:left, 3:center 4:right
# ws.Columns('B').HorizontalAlignment = 2
# ws.Columns('C').HorizontalAlignment = 2
# ws.Columns('D').HorizontalAlignment = 2
# ws.Rows(1).HorizontalAlignment = 3
# ws.Rows(1).Insert(1)  # 1행에 1줄 추가, 2줄 이상 안됨
# ws.Rows(7).Insert(1)
# ws.Rows(2).Interior.ColorIndex = 15  # 1 row 전체의 색상 변환  0:white 1:black 3:red 15: 회색
# ws.Rows(1).Interior.ColorIndex = 36  # 1 row 전체의 색상 변환  36:yellow 0:white 1:black 3:red 15: 회색
# ws.Range('C10').RowHeight = 100  # 컬럼 전체의 height 변경
# ws.Columns('A').HorizontalAlignment = 2  # 컬럼 전체 2:left, 3:center 4:right
# ws.Columns('B').HorizontalAlignment = 2
# ws.Rows(10).VerticalAlignment = 1  # 행 전체 1:top 2:center 3:bottom
# ws.Range('A8:F10').VerticalAlignment = 3  # 해당 범위 1:top 2:center 3:bottom
# ws.Range('A1').ColumnWidth = 40  # 해당 column width
# ws.Range('B1').ColumnWidth = 40
# ws.Range('C1:L1').ColumnWidth = 15
# ws.Range('A1').Value = f"작업일보 요약"
# ws.Range('A2').Value = 'Site'
# ws.Range('A1').Font.Size = 24
# ws.UsedRange()
# nCol = ws.UsedRange.Columns.Count  # column수 확인
# nRow = ws.UsedRange.Rows.Count  # row수 확인
# ws.Cells(9, nCol + 1).Value = "sum"
# ws.Range('A3:B3').Merge()
# ws.Range('A2:L3').Font.Bold = True
# ws.Range('B2:C3').Borders.Weight = 2  # 1:실선, 2:일반, 3:굵게 4:더굵게
# ws.Range('B2:C3').Borders.LineStyle = 1  # 0:없음 1:실선 2:점선 3:파선 4:
# ws.Range('B2:C3').Borders.ColorIndex = 15  # 1:black 2:white 3:red 15:gray
# ws.Range('F9').NumberFormat = '#,####'
# Common Formatting
# ws.Columns(1).Delete() # 첫째열 삭제




