import sys
sys.path.append(r"D:\Project_Python\webMDChecker\MDChecker\MailControler")

from model_mail_get import EmailAttachDownloader
from model_excel_export import SummaryExcelData
from model_excel_import import ExcelDataConverter
from model_data import *
from model_excel_styler import ExcelFormMaker
from model_mail_styler import MailFormMaker
from model_mail_send import MailSender



class PensionDataConverter:
    def __init__(self, dc, file_name, month):
        self.dc = dc
        self.file_name = file_name
        self.month = month

    def pension_data_converter(self):
        dc = self.dc
        file_name = self.file_name
        month = self.month

        # request pension data : all
        df_pension = dc.request_pension_data_without_correct_date(file_name, month)
        # 수정일을 고려하지 않고 합쳐서 가져오는 것으로 변경
        # Index(['근로년월', '현장명p', '업체명p', '공제가입번호', '확정일수', '인원수', 'sheet_name', 'attachment', 'send_date'], dtype='object')
        df_pension = df_pension.astype({'확정일수': float})
        df_pension['현장명p'] = df_pension['현장명p'].str.strip()
        df_pension['업체명'] = df_pension['업체명p'].copy()

        # DB에서 현장명p 중복 제거
        columns = ['현장명p']
        dc.drop_duplicates_from_DB_table('mdc_mst_site', columns)

        # request matching table : site code - site name
        df_mst_site = dc.call_df_from_db_with_column_name('mdc_mst_site')
        # Index(['현장코드', '현장명p'], dtype='object')
        df_mst_site['현장명p'] = df_mst_site['현장명p'].str.strip()
        df_mst_site.drop_duplicates(subset=['현장명p'], inplace=True)

        # add site_code on pension data
        df_pension = pd.merge(df_pension, df_mst_site, how='left', on='현장명p')

        # DB에서 업체명 중복 제거
        columns = ['업체명key', '업체명']
        dc.drop_duplicates_from_DB_table('mdc_mst_subcon', columns)
        # request matching table : subcon
        df_mst_subcon = dc.call_df_from_db_with_column_name('mdc_mst_subcon')
        df_mst_subcon['업체명'] = df_mst_subcon['업체명'].str.strip()
        df_mst_subcon.drop_duplicates(subset=['업체명'], inplace=True)

        # 업체명을 업체명key로 변환한 후 업체명 컬럼에 덮어씌움
        df_pension = pd.merge(df_pension, df_mst_subcon, how='left', on='업체명')
        df_pension.loc[lambda df: df['업체명key'].isna(), ['업체명key']] = ""
        df_pension['업체명'] = df_pension.apply(lambda x: x['업체명key'] if x['업체명key'] != "" else x['업체명'], axis=1)

        if df_pension['현장코드'].nunique() == 1:
            siteCode = df_pension[~df_pension['현장코드'].isna()]['현장코드'].unique()[0]
        else:
            return None, False
        df_pension.loc[lambda df: df['현장코드'].isna(), ['현장코드']] = siteCode

        send_date = max(df_pension['send_date'].str[:16])
        df_pension['send_date'] = send_date

        # slice pension data by site_code
        df_pension_temp = df_pension[df_pension['현장코드'] == siteCode]

        # month = "2022-05"
        # # method 1. 수정일 기준
        # max_date = df_pension_temp['수정일'].max()
        # a_month_before = datetime.strptime(max_date, '%Y%m%d') - relativedelta(months=1)
        # a_month_before = datetime.strftime(a_month_before, '%Y-%m')

        # # method 2. 받은달 전달 기준
        # max_date = df_pension_temp['send_date'].max()
        # a_month_before = datetime.strptime(max_date[:10], '%Y-%m-%d') - relativedelta(months=1)
        # ################# 현재 자료가 5월달이기 때문에 2개월 전으로 셋팅하나, 6월자료는 반드시 1개월 전으로 고칠 것
        # a_month_before = datetime.strftime(a_month_before, '%Y-%m')

        # request ilbo md by site_code / month
        df = dc.request_ilbo_md_by_site(siteCode, month)

        if not df.empty:
            # 업체명p를 업체명key로 변환한 후 업체명 컬럼에 덮어씌움
            df = pd.merge(df, df_mst_subcon, how='left', on='업체명')
            df.loc[lambda df: df['업체명key'].isna(), ['업체명key']] = ""
            df['업체명'] = df.apply(lambda x: x['업체명key'] if x['업체명key'] != "" else x['업체명'], axis=1)

        # add ilbo md on pension data that sliced each site_code/monthly
        df_pension_temp = df_pension_temp[['근로년월', '현장코드', 'send_date', '현장명p', '업체명', '확정일수']]
        df_pension_temp.drop_duplicates(inplace=True)
        # sum by 업체명 ignoring 현장명p(즉 다른 공구, 혹은 다른 등록번호)
        df_pension_temp_1 = df_pension_temp.groupby(['근로년월', '현장코드', 'send_date', '업체명']).sum()
        df_pension_temp_1 = df_pension_temp_1.reset_index()
        df_pension_temp_2 = df_pension_temp.drop_duplicates(['근로년월', '현장코드', 'send_date', '업체명'])
        df_pension_temp_3 = pd.merge(
            df_pension_temp_2[['근로년월', '현장코드', 'send_date', '업체명', '현장명p']],
            df_pension_temp_1,
            how='right', on=['근로년월', '현장코드', 'send_date', '업체명'])
        df = pd.merge(df, df_pension_temp_3, how='outer', on=['현장코드', '업체명'])


        # 그동안 작업했던 df_pension데이터가 작업일보와 합쳐질 때,
        # 작업일보상에만 있는 (퇴직공제부금에는 없는) 업체의 데이터를 만들어주는 작업
        df.loc[lambda df: df['근로년월'].isna(), ['근로년월']] = month
        df.loc[lambda df: df['현장명p'].isna(), ['현장명p']] = "◎ 퇴직공제부금 미등록"
        df.loc[lambda df: df['send_date'].isna(), ['send_date']] = send_date

        # format data
        df = df[['근로년월', '현장코드', '현장명p', 'send_date', '업체명', '근로자수', '소장', '직원', '확정일수']]
        df = df.sort_values(['확정일수', '근로자수'], ascending=False, kind='mergesort')
        df = df.sort_values(['근로년월', '현장코드'], ascending=True, kind='mergesort')
        df = df.rename(columns={'근로년월': '분석월','send_date': '수신일시', '근로자수': '출역일수', '소장': '소장출역', '직원': '직원출역'})
        df = df.fillna(0)

        return siteCode, df


if __name__ == '__main__':
