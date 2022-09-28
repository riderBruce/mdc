from model_data import *
import os

class Data_extractor:
    def __init__(self, dc, siteCode=None):
        self.dc = dc
        if siteCode:
            self.siteCode = siteCode
        else:
            self.siteCode = "PLNT"

    def name_output_excel_file(self):
        # 상위폴더까지 생성 / 폴더 존재시 에러 없음
        os.makedirs(resultPath, exist_ok=True)
        # set fileName
        outputFileName = resultPath + f"\DailyReport({self.siteCode})_{sRunTime}.xlsx"
        if os.path.exists(outputFileName):
            os.remove(outputFileName)
        return outputFileName

    def request_table_summary_sheet16(self, siteCode=None):
        """Raw all data"""
        curs = self.dc.conn.cursor()
        sSql = f"select reportdate, projectcode, discipline, subcontractor, contractnumber, " \
               f"cwa, cwp, iwp, browngreenfield, category, tagnostrcturenopipingmaterial, " \
               f"activity, level1, level2, level3, " \
               f"sum(amount) " \
               f"from job_plnt_data " \
               f"group by reportdate, projectcode, discipline, subcontractor, contractnumber, " \
               f"cwa, cwp, iwp, browngreenfield, category, tagnostrcturenopipingmaterial, " \
               f"activity, level1, level2, level3;"
        curs.execute(sSql)
        data = curs.fetchall()
        column_names = ['reportdate', 'project', 'discipline', 'subcontractor', 'contractnumber',
                        'cwa', 'cwp', 'iwp', 'browngreenfield', 'category', 'No',
                        'activity', 'level1', 'level2', 'level3', 'Amount']
        df = pd.DataFrame(data, columns=column_names)
        df = df.astype({
            'project':str, 'discipline':str, 'subcontractor':str, 'contractnumber':str,
            'cwa':str, 'cwp':str, 'iwp':str, 'browngreenfield':str, 'category':str, 'No':str,
            'activity':str, 'level1':str, 'level2':str, 'level3':str,
            'Amount': int
        })
        df.replace('NaN', '', inplace=True)
        return df

    def consist_excel_data(self, outputFileName):
        # Raw
        df = self.request_table_summary_sheet16(self.siteCode)
        df = df.pivot_table(values='Amount',
                            index=['project', 'discipline', 'subcontractor', 'contractnumber',
                                   'cwa', 'cwp', 'iwp', 'browngreenfield', 'category', 'No',
                                   'activity', 'level1', 'level2', 'level3'],
                            columns='reportdate', aggfunc='sum', fill_value="", dropna=True)
        df = df.reset_index()
        Raw = df.copy()

        with pd.ExcelWriter(outputFileName) as writer:
            if len(Raw) > 0:
                Raw.to_excel(writer, sheet_name="Raw")



if __name__ == '__main__':
    dc = DataControl()
    dc.db_host = '10.171.94.66'
    dc.db_name = 'JobReport'
    dc.db_user = 'postgres'
    dc.db_pwd = 'nam1004'
    conn = pg2.connect(
        'host={0} dbname={1} user={2} password={3}'.format(dc.db_host, dc.db_name, dc.db_user,
                                                           dc.db_pwd))
    dc.conn = conn
    data_extractor = Data_extractor(dc)
    outputFileName = data_extractor.name_output_excel_file()
    data_extractor.consist_excel_data(outputFileName)
    os.startfile(outputFileName)

