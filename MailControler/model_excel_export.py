import numpy as np
import win32com.client
import re
import os

from model_data import *

class SummaryExcelData:
    def __init__(self, dc, siteCode, month=None):
        self.dc = dc
        self.siteCode = siteCode
        if month:
            self.month = month
        else:
            self.month = "ALL"

    def name_output_excel_file(self):
        # 상위폴더까지 생성 / 폴더 존재시 에러 없음
        os.makedirs(resultPath, exist_ok=True)
        # set fileName
        outputFileName = resultPath + f"\\Reply({self.siteCode}_{self.month})_{sRunTime}.xlsx"
        if os.path.exists(outputFileName):
            os.remove(outputFileName)
        return outputFileName

    def excel_writer(self, outputFileName):
        df_1 = self.excel_summary_page()
        # 누계표 - 금번 첫 현장발송시(9/23 예정)에는 누계표 제외하고 비교표만 보낼 예정
        # df_2 = self.excel_accumulate_page()

        # create excel_file
        with pd.ExcelWriter(outputFileName) as writer:
            if not df_1.empty:
                df_1.to_excel(writer, sheet_name="비교표")
            # if not df_2.empty:
            #     df_2.to_excel(writer, sheet_name="누계표")

        logWrite(f'[파일생성 complete  ] : {self.siteCode} - {outputFileName}')

    def excel_summary_page(self):
        dc = self.dc
        df = dc.request_table_result_by_site(self.siteCode, self.month)
        df['구분'] = df['업체명'].apply(lambda x: '당사' if '현대건설' in x else '협력업체')
        df = df.sort_values(['구분', '업체명', '확정일수'], ascending=[True, True, False])

        if '협력업체' not in df['구분'].unique():
            df_temp = pd.DataFrame([['협력업체', '', 0, 0, 0, 0]], columns=['구분', '업체명', '출역일수', '소장출역', '직원출역', '확정일수'])
            df = pd.concat([df, df_temp])
        if '당사' not in df['구분'].unique():
            df_temp = pd.DataFrame([['당사', '현대건설(주)', 0, 0, 0, 0]], columns=['구분', '업체명', '출역일수', '소장출역', '직원출역', '확정일수'])
            df = pd.concat([df, df_temp])

        df = df.set_index(['구분', '업체명'])
        temp = [d for k, d in df.groupby(level=0)]
        df_hd = temp[0]
        df_subcon = temp[1]
        df_subcon = df_subcon.append(df_subcon.sum().rename(('협력업체', '소계')))
        df_result = pd.concat([df_hd, df_subcon]).append(df.sum().rename(('현장총계', '')))

        # add rate column
        df_result['대비'] = df_result.apply(lambda x: round(x.확정일수 / x.출역일수, 2) if x.출역일수 and x.확정일수 else 0, axis=1)
        # df_result['대비'] = df_result.apply(lambda x: str(round(x.확정일수 / x.출역일수 * 100))+"%" if x.출역일수 and x.확정일수 else "", axis=1)

        # 현장명 : 현장코드 + 현장명
        site_name = dc.request_site_name_by_site_code(self.siteCode)
        df_result.loc[:, '현장명'] = f'[{self.siteCode}]\n{site_name}'

        # define column type
        df_result = df_result.reset_index()
        df_result['비고'] = df_result.apply(lambda x: "◎ 전자카드 근무관리시스템 미등록" if (
                x['확정일수'] == 0 and
                x['업체명'] not in ('현대건설(주)', '', '소계')) else "", axis=1)
        df_result = df_result[['현장명', '구분', '업체명', '출역일수', '소장출역', '직원출역', '확정일수', '대비', '비고']]
        df_result = df_result.astype({
            '현장명': str, '구분': str, '업체명': str,
            '출역일수': int, '소장출역': int, '직원출역': int, '확정일수': float,
            '대비': float, '비고': str
        })

        return df_result

    def excel_accumulate_page(self):
        dc = self.dc
        df = dc.request_accumulate_data_by_site(self.siteCode, self.month)
        # columns = ['업체명', '분석월', '작업일보', '퇴직공제부금']
        df = df.astype({'작업일보': float, '퇴직공제부금': float})

        # 협력업체 / 당사가 없는 경우
        df['구분'] = df['업체명'].apply(lambda x: '당사' if '현대건설' in x else '협력업체')
        if '협력업체' not in df['구분'].unique():
            df_temp = pd.DataFrame([['협력업체', '', self.month, 0, 0]], columns=['구분', '업체명', '분석월', '작업일보', '퇴직공제부금'])
            df = pd.concat([df, df_temp])
        if '당사' not in df['구분'].unique():
            df_temp = pd.DataFrame([['당사', '현대건설(주)', self.month, 0, 0]], columns=['구분', '업체명', '분석월', '작업일보', '퇴직공제부금'])
            df = pd.concat([df, df_temp])

        # 자료가 없는 달의 컬럼 생성
        M0 = self.month
        M1 = (datetime.strptime(M0, '%Y-%m') - relativedelta(months=1)).strftime('%Y-%m')
        M2 = (datetime.strptime(M0, '%Y-%m') - relativedelta(months=2)).strftime('%Y-%m')
        M3 = (datetime.strptime(M0, '%Y-%m') - relativedelta(months=3)).strftime('%Y-%m')

        for m in [M0, M1, M2, M3]:
            if m not in df['분석월'].unique():
                df_temp = df.drop_duplicates('업체명').copy()
                df_temp.loc[:, '분석월'] = m
                df_temp.loc[:, '작업일보'] = 0
                df_temp.loc[:, '퇴직공제부금'] = 0
                df = pd.concat([df, df_temp])

        df = df[['구분', '업체명', '분석월', '작업일보', '퇴직공제부금']]
        df = df.sort_values(['구분', '업체명', '분석월'], ascending=[True, True, True])
        df = df.set_index(['구분', '업체명', '분석월'])
        df = df.stack(-1)
        df = df.unstack(-2)
        df = df.fillna(0)

        temp = [d for k, d in df.groupby(level=0)]
        df_hd, df_subcon = temp

        # 당사 누계
        df_hd_cumsum = df_hd.sum(level=-1).cumsum(axis=1)
        df_hd_cumsum = df_hd_cumsum.rename({'작업일보': ('당사', '누 계', '작업일보'), '퇴직공제부금': ('당사', '누 계', '퇴직공제부금')})
        df_hd = pd.concat([df_hd, df_hd_cumsum])

        # 협력업체 소계 및 누계
        df_subcon_sum = df_subcon.sum(level=-1)
        df_subcon_sum = df_subcon_sum.rename({'작업일보': ('협력업체', '소 계', '작업일보'), '퇴직공제부금': ('협력업체', '소 계', '퇴직공제부금')})
        df_subcon_cumsum = df_subcon.sum(level=-1).cumsum(axis=1)
        df_subcon_cumsum = df_subcon_cumsum.rename({'작업일보': ('협력업체', '누 계', '작업일보'), '퇴직공제부금': ('협력업체', '누 계', '퇴직공제부금')})
        df_subcon = pd.concat([df_subcon, df_subcon_sum])
        df_subcon = pd.concat([df_subcon, df_subcon_cumsum])

        # 합치기
        df_result = pd.concat([df_hd, df_subcon])

        # 전체 합계 및 누계
        df_sum = df.sum(level=-1)
        df_sum = df_sum.rename({'작업일보': ('현장총계', '', '작업일보'), '퇴직공제부금': ('현장총계', '', '퇴직공제부금')})
        df_cumsum = df.sum(level=-1).cumsum(axis=1)
        df_cumsum = df_cumsum.rename({'작업일보': ('현장누계', '', '작업일보'), '퇴직공제부금': ('현장누계', '', '퇴직공제부금')})
        df_result = pd.concat([df_result, df_sum])
        df_result = pd.concat([df_result, df_cumsum])

        # 비고 삽입
        df_result.loc[:, '비 고'] = ''
        return df_result



