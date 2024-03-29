from sqlalchemy import create_engine
import psycopg2 as pg2
import pandas as pd
import numpy as np
import psutil
import os
import shutil
from pathlib import Path

from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta

# root directory
rootPath = r'D:\Project_Data\webMDChecker'
logFile = f"{rootPath}\Log\Log_{datetime.now().strftime('%Y-%m-%d')}_MD.log"
savePath = f"{rootPath}\Files_MD\{datetime.now().strftime('%Y-%m')}"   # 첨부파일 저장위치
resultPath = f"{rootPath}\Result_"

# 상위폴더까지 생성 / 폴더 존재시 에러 없음
os.makedirs(f"{rootPath}\Log", exist_ok=True)
os.makedirs(savePath, exist_ok=True)

# set run time
sRunTime = datetime.now().strftime('%Y-%m-%d_%H%M')
T_Date = datetime.now().strftime('%Y-%m-%d')                         # Today
D1_Day = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')   # D-1 날짜
D2_Day = (datetime.now() - timedelta(days=2)).strftime('%Y-%m-%d')   # D-2 날짜
D3_Day = (datetime.now() - timedelta(days=3)).strftime('%Y-%m-%d')   # D-3 날짜
if int(T_Date[8:10]) < 10:
    F_Date = (datetime.now() - timedelta(days=10)).strftime('%Y-%m') + '-01'  # 집계기간 From / 월초(매달 10일 이전)
else:
    F_Date = datetime.now().strftime('%Y-%m') + '-01'  # 집계기간 From / 월중~월말


def logWrite(logStr):
    print(logStr)
    f = open(logFile, 'a', encoding='UTF-8')
    f.write(datetime.now().strftime('%Y-%m-%d %H%M%S') + " : " + logStr + "\n")
    f.close()

def kill_excel():
    for proc in psutil.process_iter():
        if proc.name() == "EXCEL.EXE":
            proc.kill()

class DataControl:
    def __init__(self, DB):
        self.sRunType = DB
        # DB 서버 정보 ###########################
        if self.sRunType == 'SERVER':

        else:



        conn = pg2.connect(
            'host={0} dbname={1} user={2} password={3}'.format(self.db_host, self.db_name, self.db_user,
                                                               self.db_pwd))
        self.conn = conn


    def call_df_from_db_with_column_name(self, db_table_name):
        # get DB data
        curs = self.conn.cursor()
        sSql = f"select * from {db_table_name};"  # order by col_01 desc
        curs.execute(sSql)
        data_all = curs.fetchall()

        # get DB name
        sSql = f"SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '{db_table_name}' order by ordinal_position;"
        curs.execute(sSql)
        column_names = curs.fetchall()

        # 가져온 컬럼명은 (이름,  ) 이런 형태이므로 첫번째 컬럼만 별도로 리스트
        column_names = [i[0] for i in column_names]
        df = pd.DataFrame(data_all, columns=column_names)

        return df

    def get_column_names(self, db_table_name):
        curs = self.conn.cursor()
        # get DB name
        sSql = f"SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '{db_table_name}' order by ordinal_position;"
        curs.execute(sSql)
        column_names = curs.fetchall()
        # 가져온 컬럼명은 (이름,  ) 이런 형태이므로 첫번째 컬럼만 별도로 리스트
        column_names = [i[0] for i in column_names]
        return column_names

    def upload_df_to_db_at_once(self, df, db_table_name):
        engine = create_engine(self.connStr)
        df.to_sql(db_table_name, engine, if_exists='append')

    def is_new_mail(self, mailStr):
        curs = self.conn.cursor()
        sSql = f"select count(*) from mdc_mails " \
               f"   where mail = '{mailStr}' and (처리현황 = '담당자확인중' or 처리현황 is null);"
        curs.execute(sSql)
        data = curs.fetchone()
        count = data[0]
        if count > 0:
            logWrite(f'[▷ 진행중           ] : {mailStr}, 이미 수신한 메일이지만 미처리 내역이 남아 재수신합니다.')
            return True
        else:
            sSql = f"select count(*) from mdc_mails " \
                   f"   where mail = '{mailStr}';"
            curs.execute(sSql)
            data = curs.fetchone()
            count = data[0]
            if count == 0:
                return True
            else:
                # logWrite(f'[▷ fail           ] : {mailStr}, 신규메일이 아닙니다.')
                return False

    def is_already_finished(self, mailStr, sendDate, senderTxt, subject, save_date, filename_stem):
        curs = self.conn.cursor()
        sSql = f"select count(*) from mdc_mails " \
               f"   where mail = '{mailStr}' and send_date = '{sendDate}' and sender = '{senderTxt}' " \
               f"       and subject = '{subject}' and attachment = '{filename_stem}' " \
               f"       and 처리현황 = '메일송부완료' " \
               f"       and save_date <> '{save_date}'; "
               # f"       and 처리현황 <> '담당자확인중' " \
               # f"       and 처리현황 is not null;"
        curs.execute(sSql)
        data = curs.fetchone()
        if data[0]:
            return True
        else:
            return False

    def insert_mail_get_list(self, mailStr, sendDate, senderTxt, subject, save_date, filename=None):
        curs = self.conn.cursor()
        sSql = f"insert into " \
               f"   mdc_mails " \
               f"   (mail, send_date, sender, subject, save_date, attachment) " \
               f"   values" \
               f"   ('{mailStr}','{sendDate}','{senderTxt}','{subject}', '{save_date}', '{filename}');"
        curs.execute(sSql)

    def drop_duplicates_by_date_from_raw(self, numbers, target_month):
        # 공제가입번호와 근로년월이 같은 첨부파일이 있는 경우 삭제
        curs = self.conn.cursor()

        if len(numbers) < 2:
            param = f"공제가입번호 = '{numbers[0]}' "
        else:
            param = f'공제가입번호 in {tuple(numbers)} '

        sSql = f"delete from mdc_raw_md " \
               f"   where attachment in (" \
               f"       select distinct attachment " \
               f"           from mdc_raw_md " \
               f"           where {param} " \
               f"               and 근로년월 = '{target_month}'" \
               f"   ) ;"
        curs.execute(sSql)

    def update_mails_sheetName(self, fileName_stem, sheetName, sendDate):
        curs = self.conn.cursor()
        sSql = f"update mdc_mails set " \
               f"   sheet_name = '{sheetName}' " \
               f"   where attachment = '{fileName_stem}' and send_date like '{sendDate[:7]}%';"
        curs.execute(sSql)

    def update_mails_process_status(self, fileName_stem, sendDate, senderTxt, text):
        curs = self.conn.cursor()
        sSql = f"update mdc_mails set " \
               f"   처리현황 = '{text}', 처리일 = '{datetime.now().strftime('%Y-%m-%d %H:%M:%S %z')}'  " \
               f"   where attachment = '{fileName_stem}' and send_date = '{sendDate}' and sender = '{senderTxt}';"
        curs.execute(sSql)

    def update_mails_process_status_by_mailStr(self, mailStr, sendDate, senderTxt, save_date, text):
        curs = self.conn.cursor()
        sSql = f"update mdc_mails set " \
               f"   처리현황 = '{text}', 처리일 = '{datetime.now().strftime('%Y-%m-%d %H:%M:%S %z')}'  " \
               f"   where mail = '{mailStr}' and send_date = '{sendDate}' and sender = '{senderTxt}' and save_date = '{save_date}';"
        curs.execute(sSql)

    def update_mails_process_status_by_siteCode(self, site_code, save_date, text):
        curs = self.conn.cursor()
        sSql = f"update mdc_mails set " \
               f"   처리현황 = '{text}', 처리일 = '{datetime.now().strftime('%Y-%m-%d %H:%M:%S %z')}' " \
               f"   where 현장코드 = '{site_code}' and save_date = '{save_date}' and 처리일 is null;"
        curs.execute(sSql)

    def update_mails_process_status_to_ongoing(self, fileName_stem, send_date, text):
        curs = self.conn.cursor()
        sSql = f"update mdc_mails set " \
               f"   처리현황 = '{text}', 처리일 = '{datetime.now().strftime('%Y-%m-%d %H:%M:%S %z')}' " \
               f"   where 현장코드 = 'None' and attachment = '{fileName_stem}'and send_date = '{send_date}' and 처리일 is null;"
        curs.execute(sSql)

    def update_mails_older_process_status_by_siteCode(self, fileName_stem, send_date, senderTxt, siteCode):
        curs = self.conn.cursor()
        sSql = f"update mdc_mails set " \
               f"   현장코드 = '{siteCode}', 처리현황 = '담당자확인/처리완료' " \
               f"   where attachment = '{fileName_stem}' and send_date = '{send_date}' and sender = '{senderTxt}' " \
               f"       and (처리현황 != '메일송부완료' or 처리현황 is null);"
        curs.execute(sSql)

    def update_mails_add_site_code(self, fileName_stem, sendDate, senderTxt):
        curs = self.conn.cursor()
        sSql = f"with T(현장코드, attachment) as (" \
               f"    select distinct 현장코드, attachment " \
               f"           from mdc_result a, mdc_raw_md b " \
               f"           where a.현장명p = b.현장명p " \
               f"               and a.현장코드 <> '코드누락'" \
               f") " \
               f"update mdc_mails c " \
               f"   set 현장코드 = T.현장코드 " \
               f"   from T" \
               f"   where c.attachment = T.attachment " \
               f"       and c.attachment = '{fileName_stem}' " \
               f"       and send_date = '{sendDate}' " \
               f"       and sender = '{senderTxt}'" \
               f"       and (c.현장코드 is null or c.현장코드 = 'None') ;"
        # sSql = f"update mdc_mails set 현장코드 = d.현장코드 " \
        #        f"   from (select distinct c.attachment, a.현장코드 " \
        #        f"           from mdc_result a, mdc_raw_md b, mdc_mails c " \
        #        f"           where a.현장명p = b.현장명p and b.attachment = c.attachment " \
        #        f"                 and a.현장코드 <> '코드누락') as d " \
        #        f"   where mdc_mails.attachment = d.attachment and mdc_mails.현장코드 is null " \
        #        f"       and mdc_mails.attachment = '{fileName_stem}' and send_date = '{sendDate}' and sender = '{senderTxt}';"
        curs.execute(sSql)

    def insert_new_mail_sender_address(self, siteCode, senderTxt):
        curs = self.conn.cursor()
        senderTxt = senderTxt.replace('<', ' ')
        senderTxt = senderTxt.replace('>', '')
        sender_data = senderTxt.split()
        sender_name = sender_data[0]
        sender_name = sender_name[:10]
        for i in sender_data:
            if '@' in i:
                sender_email = i
                break
        else:
            sender_email = None
            return logWrite(f'[이메일 주소 자동 등록 ] : {siteCode}/{senderTxt} 자동 등록 실패')

        if sender_name in ('최다희', '김영일', '남효정'):
            return
        # sSql = f"delete from mdc_address where email = '{sender_email}';"
        # cur.execute(sSql)


        sSql = f"insert into mdc_address(이름, 메일주소, 현장코드) values('{sender_name}', '{sender_email}', '{siteCode}');"
        try:
            curs.execute(sSql)
            return logWrite(f'[이메일 주소 자동 등록 ] : {siteCode}/{senderTxt} 자동 등록 완료')
        except Exception as ex:
            return logWrite(f'[이메일 주소 자동 등록 ] : {siteCode}/{senderTxt} 자동 등록 실패 \n {ex} \n{sSql}')

    def drop_duplicates_from_DB_table(self, table_name, columns):
        curs = self.conn.cursor()
        column_names_to_text = ', '.join(columns)
        sSql = f"WITH T AS " \
               f"( " \
               f"   SELECT  *, ctid, row_number() OVER (PARTITION BY {column_names_to_text} ORDER BY ctid) " \
               f"   FROM {table_name}" \
               f") " \
               f"DELETE " \
               f"   FROM {table_name} " \
               f"   WHERE ctid IN  " \
               f"   (" \
               f"       SELECT ctid  " \
               f"           FROM T " \
               f"           WHERE row_number >= 2" \
               f"   ) ;"
        try:
            curs.execute(sSql)
            # print(table_name, ": 중복 삭제 완료")
        except Exception as ex:
            print(sSql, ex)

    def insert_new_site_name_to_mdc_mst_site(self, df):
        curs = self.conn.cursor()
        df = df[['현장코드', '현장명p']]
        df = df[~df['현장명p'].str.contains('◎ 퇴직공제부금 미등록')]
        df = df.drop_duplicates()
        data = df.values
        sSql = f"delete from mdc_mst_site " \
               f"   where 현장코드=%s and 현장명p=%s; "
        curs.executemany(sSql, data)
        sSql = f"insert into mdc_mst_site " \
               f"   (현장코드, 현장명p) " \
               f"   values(%s, %s); "
        curs.executemany(sSql, data)
        self.conn.commit()

    def update_mails_add_site_code_None(self, fileName_stem, sendDate):
        curs = self.conn.cursor()
        sSql = f"update mdc_mails set " \
               f"   현장코드 = 'None' " \
               f"   where attachment = '{fileName_stem}' and send_date like '{sendDate[:7]}%';"
        curs.execute(sSql)

    def update_attachment_list(self, projectCode, subcontractor, discipline, reportDate, sRunTime,row_File):
        curs = self.conn.cursor()
        sSql = f"update job_attach set 현장코드 = '{projectCode}', 업체 = '{subcontractor} / {discipline}', 날짜 = '{reportDate}', 집계일시 = '{sRunTime}' " \
               f"where 첨부파일 = '{row_File}'"
        curs.execute(sSql)

    def delete_not_attached_list(self):
        curs = self.conn.cursor()
        sSql = f"delete from job_attach where 날짜 is null;"
        curs.execute(sSql)

    def upload_ilbo_df_to_DB(self, df, projectCode, reportDate, subcontractor, discipline, dbTable):
        curs = self.conn.cursor()
        # delete old data
        sSql = f"delete from {dbTable} where reportdate='{reportDate}' and subcontractor='{subcontractor}' and projectcode='{projectCode}' and discipline = '{discipline}';"
        try:
            curs.execute(sSql)
        except Exception as ex:
            print(ex, sSql)
        # upload to DB
        sColumnName = ','.join(df.columns.tolist())
        for index, row in df.iterrows():
            sSql = f"insert into {dbTable}({sColumnName}) values({'%s,' * (len(row) - 1)}%s)"
            try:
                curs.execute(sSql, list(row))
            except Exception as ex:
                print(ex, sSql)
                return
        logWrite(f'[▶ finish         ] : 작업일보 DB 업로드 완료, {projectCode}, {reportDate}, {subcontractor}, {discipline}')
        return

    def request_table_in_mail_body(self, siteCode):
        if siteCode == "PLNT":
            pQuery_siteCode = "and 현장코드 in (select distinct(projectcode) from job_plnt_data) "
        else:
            pQuery_siteCode = f"and 현장코드 = '{siteCode}' "
        curs = self.conn.cursor()
        sSql = f"select 현장코드, 업체, 발송자, 날짜, 발송일시, 집계일시  " \
               f"from job_attach " \
               f"where 발송일시 >= '{D3_Day}' {pQuery_siteCode} " \
               f"order by cast(집계일시 as date) desc, 현장코드; "
        curs.execute(sSql)
        data = curs.fetchall()
        column_names = ['현장코드', '업체', '발송자', '날짜', '발송일시', '집계일시']
        df = pd.DataFrame(data, columns=column_names)
        try:
            summary_time = datetime.strptime(D1_Day + ' 05:00:00 +0900', '%Y-%m-%d %H:%M:%S %z')
            df = df[df['발송일시'].apply(lambda x: datetime.strptime(x, '%Y-%m-%d %H:%M:%S %z') >= summary_time)]
            logWrite(f"{summary_time} 이후에 수신된 메일만 첨부에 넣습니다.")
        except:
            pass
        return df

    def request_local_timezone(self, siteCode):
        curs = self.conn.cursor()
        sSql = f"select 타임존 from job_mst_site where 현장코드 = '{siteCode}'"
        curs.execute(sSql)
        timezone = curs.fetchall()
        return timezone[0][0]

    def delete_old_data(self, dbTable, param=None):
        curs = self.conn.cursor()
        sSql = f"delete from {dbTable} {param};"
        try:
            curs.execute(sSql)
        except Exception as ex:
            print(ex, sSql)
            self.conn.rollback()
            return ex
        return True

    def insert_data_to_db(self, df, dbTable):
        curs = self.conn.cursor()
        sColumnName = ','.join(df.columns.tolist())
        for index, row in df.iterrows():
            sSql = f"insert into {dbTable}({sColumnName}) values({'%s,' * (len(row) - 1)}%s)"
            try:
                values = list(row)
                curs.execute(sSql, values)
            except Exception as ex:
                print(ex, sSql)
                self.conn.rollback()
                return ex
        return True

    def insert_subcon_maching_data_to_db(self, subcon_name_key, subcon_name_simular):
        curs = self.conn.cursor()
        sSql = f"delete from mdc_mst_subcon where 업체명key = '{subcon_name_key}' and 업체명 = '{subcon_name_simular}';"
        try:
            curs.execute(sSql)
        except:
            return False
        sSql = f"insert into mdc_mst_subcon(업체명key, 업체명) values('{subcon_name_key}', '{subcon_name_simular}');"
        try:
            curs.execute(sSql)
            self.conn.commit()
        except:
            return False
        return True

    def insert_address_data_to_db(self, address_name, address_mail, address_site_code, address_department, address_managing_bonbu):
        curs = self.conn.cursor()
        # sSql = f"delete from mdc_address where 메일주소 = '{address_mail}' ;"
        # try:
        #     curs.execute(sSql)
        # except:
        #     return False
        sSql = f"insert into mdc_address(이름, 메일주소, 현장코드, 부서, 담당본부) " \
               f"   values('{address_name}', '{address_mail}', '{address_site_code}', '{address_department}', '{address_managing_bonbu}');"
        try:
            curs.execute(sSql)
            self.conn.commit()
        except:
            return False
        return True

    def delete_subcon_maching_data_to_db(self, subcon_name_key, subcon_name_simular):
        curs = self.conn.cursor()
        sSql = f"delete from mdc_mst_subcon where 업체명key = '{subcon_name_key}' and 업체명 = '{subcon_name_simular}';"
        try:
            curs.execute(sSql)
            self.conn.commit()
        except:
            return False
        return True

    def delete_address_data_to_db(self, del_address):
        curs = self.conn.cursor()
        sSql = f"delete from mdc_address where 메일주소 = '{del_address}' ;"
        try:
            curs.execute(sSql)
            self.conn.commit()
        except:
            return False
        return True

    def request_ilbo_md_by_site(self, siteCode, month):
        curs = self.conn.cursor()
        # 전체 출역수
        sSql = f"select 업체, sum(인원) from " \
               f" (select 현장코드, 현장명, 업체, 날짜, 대공종, 공종, 직종, 인원  " \
               f"   from job_data  " \
               f"   where 현장코드 = '{siteCode}' and 날짜 like '{month}%' and 인원 <> 0 " \
               f"         and 직종 <> '소장(협력사)'  and 직종 <> '직원(협력사)' and 직종 <> '직원'" \
               f"   union all " \
               f"       select 현장코드, 현장명, 업체, 날짜, 대공종, 공종, 직종, 인원  " \
               f"           from job_data_hpms  " \
               f"           where 현장코드 = '{siteCode}' and 날짜 like '{month}%' and 인원 <> 0 " \
               f"                and 직종 <> '소장(협력사)'  and 직종 <> '직원(협력사)' and 직종 <> '직원'" \
               f") a " \
               f" group by 업체;"
        curs.execute(sSql)
        data = curs.fetchall()
        column_names = ["업체", "근로자수"]
        df = pd.DataFrame(data, columns=column_names)

        # 소장 출역수
        sSql = f"select 업체, sum(인원) from " \
               f" (select 현장코드, 현장명, 업체, 날짜, 대공종, 공종, 직종, 인원  " \
               f"   from job_data  " \
               f"   where 현장코드 = '{siteCode}' and 날짜 like '{month}%' and 인원 <> 0 " \
               f"         and 직종 = '소장(협력사)' " \
               f"   union all " \
               f"       select 현장코드, 현장명, 업체, 날짜, 대공종, 공종, 직종, 인원  " \
               f"           from job_data_hpms  " \
               f"           where 현장코드 = '{siteCode}' and 날짜 like '{month}%' and 인원 <> 0 " \
               f"                and 직종 = '소장(협력사)' " \
               f") a " \
               f" group by 업체;"
        curs.execute(sSql)
        data = curs.fetchall()
        column_names = ["업체", "소장"]
        df1 = pd.DataFrame(data, columns=column_names)

        # 직원 출역수
        sSql = f"select 업체, sum(인원) from " \
               f" (select 현장코드, 현장명, 업체, 날짜, 대공종, 공종, 직종, 인원  " \
               f"   from job_data  " \
               f"   where 현장코드 = '{siteCode}' and 날짜 like '{month}%' and 인원 <> 0 " \
               f"         and 직종 in ('직원(협력사)', '직원') " \
               f"   union all " \
               f"       select 현장코드, 현장명, 업체, 날짜, 대공종, 공종, 직종, 인원  " \
               f"           from job_data_hpms  " \
               f"           where 현장코드 = '{siteCode}' and 날짜 like '{month}%' and 인원 <> 0 " \
               f"                and 직종 in ('직원(협력사)', '직원') " \
               f") a " \
               f" group by 업체;"
        curs.execute(sSql)
        data = curs.fetchall()
        column_names = ["업체", "직원"]
        df2 = pd.DataFrame(data, columns=column_names)

        df = pd.merge(df, df1, how='outer', on='업체')
        df = pd.merge(df, df2, how='outer', on='업체')
        df['현장코드'] = siteCode
        df = df.rename(columns={'업체': '업체명'})
        df = df[['현장코드', '업체명', '근로자수', '소장', '직원']]
        return df

    def request_table_result_by_site(self, site_code, month):
        curs = self.conn.cursor()
        sSql = f"select 업체명, sum(출역일수), sum(소장출역), sum(직원출역), sum(확정일수) " \
               f"   from mdc_result " \
               f"   where 현장코드 = '{site_code}' and 분석월 = '{month}' " \
               f"   group by 업체명 ;"
        curs.execute(sSql)
        data = curs.fetchall()
        if '현대건설(주)' not in [i[0] for i in data]:
            data.append(('현대건설(주)', 0, 0, 0, 0))
        columns = ['업체명', '출역일수', '소장출역', '직원출역', '확정일수']
        df = pd.DataFrame(data, columns=columns)
        return df

    def request_accumulate_data_by_site(self, site_code, month):
        curs = self.conn.cursor()
        month_3_month = (datetime.strptime(month, '%Y-%m') - relativedelta(months=3)).strftime('%Y-%m')
        sSql = f"select 업체명, 분석월, sum(출역일수) as 작업일보, sum(확정일수) as 퇴직공제부금 " \
               f"   from mdc_result " \
               f"   where 현장코드 = '{site_code}' and 분석월 >= '{month_3_month}' and  분석월 <= '{month}' " \
               f"   group by 업체명, 분석월 ;"
        curs.execute(sSql)
        data = curs.fetchall()
        if '현대건설(주)' not in [i[0] for i in data]:
            months = list(dict.fromkeys([i[1] for i in data]))
            for m in months:
                data.append(('현대건설(주)', m, 0, 0))
        columns = ['업체명', '분석월', '작업일보', '퇴직공제부금']
        df = pd.DataFrame(data, columns=columns)
        return df

    def request_table_result_by_site_temp(self, site_code, month):
        curs = self.conn.cursor()
        sSql = f"select 현장명p, 업체명, 출역일수, 소장출역, 직원출역, 확정일수 " \
               f"   from mdc_result " \
               f"   where 현장코드 = '{site_code}' and 분석월 = '{month}';"
        curs.execute(sSql)
        data = curs.fetchall()
        columns = ['현장명p', '업체명', '출역일수', '소장출역', '직원출역', '확정일수']
        df = pd.DataFrame(data, columns=columns)
        return df

    def request_site_name_by_site_code(self, site_code):
        curs = self.conn.cursor()
        sSql = f"select distinct 현장명 from job_data where 현장코드 = '{site_code}' limit 1;"
        curs.execute(sSql)
        data = curs.fetchone()
        if data is None:
            sSql = f"select distinct 현장명 from job_data_hpms where 현장코드 = '{site_code}';"
            curs.execute(sSql)
            data = curs.fetchone()
            if data is None:
                sSql = f"select distinct col_04 from prd_rawdata_md where col_03 = '{site_code}';"
                curs.execute(sSql)
                data = curs.fetchone()
                if data is None:
                    sSql = f"select distinct left(현장명p, 10) from mdc_mst_site where 현장코드 = '{site_code}' limit 1;"
                    curs.execute(sSql)
                    data = curs.fetchone()
                    if data is None:
                        site_name = "미등록 현장명"
                        return site_name
        site_name = data[0]
        return site_name

    def request_pension_data_without_correct_date(self, attachment, month):
        curs = self.conn.cursor()
        sSql = f"select 근로년월, 현장명p, 업체명p, 공제가입번호, sum(확정일수), sum(인원수), sheet_name, attachment, send_date " \
               f"   from mdc_raw_md " \
               f"   where attachment = '{attachment}' and 근로년월 = '{month}' " \
               f"   group by 근로년월, 현장명p, 업체명p, 공제가입번호, sheet_name, attachment, send_date;"
        curs.execute(sSql)
        data = curs.fetchall()
        if not data:
            return None
        columns = ['근로년월', '현장명p', '업체명p', '공제가입번호', '확정일수', '인원수', 'sheet_name', 'attachment', 'send_date']
        df = pd.DataFrame(data, columns=columns)
        df = df.sort_values(['근로년월', '업체명p', '확정일수'], ascending=[False, True, False])
        return df

    def request_pension_data(self, attachment):
        curs = self.conn.cursor()
        sSql = f"select 현장명p, 업체명p, 공제가입번호, 수정일, sum(확정일수), sum(인원수), sheet_name, attachment, send_date " \
               f"   from mdc_raw_md " \
               f"   where attachment = '{attachment}'" \
               f"   group by 현장명p, 업체명p, 공제가입번호, 수정일, sheet_name, attachment, send_date;"
        curs.execute(sSql)
        data = curs.fetchall()
        if not data:
            return None
        columns = ['현장명p', '업체명p', '공제가입번호', '수정일', '확정일수', '인원수', 'sheet_name', 'attachment', 'send_date']
        df = pd.DataFrame(data, columns=columns)
        df = df.sort_values(['업체명p', '확정일수'], ascending=[True, False])
        return df

    def request_address_mailto(self, site_code):
        curs = self.conn.cursor()
        sSql = f"select 메일주소 from mdc_address where 현장코드 = '{site_code}';"
        curs.execute(sSql)
        data = curs.fetchall()
        add_list = list(dict.fromkeys([i[0] for i in data]))
        mailtoList = ', '.join(add_list)
        # logWrite(f"request address for site --- 현장코드 : {site_code}, 담당자수 : {len(data)}, 메일주소 : {mailtoList}")
        return mailtoList


    def request_address_cc(self):
        curs = self.conn.cursor()
        sSql = f"select 메일주소 from mdc_address where 부서 = 'RM';"
        curs.execute(sSql)
        data = curs.fetchall()
        add_list = list(dict.fromkeys([i[0] for i in data]))
        mailtoCcList = ', '.join(add_list)
        # logWrite(f"request address cc --- 담당자수 : {len(data)}, 메일주소 : {mailtoList}")
        return mailtoCcList

    def request_address_bcc(self):
        curs = self.conn.cursor()
        sSql = f"select 메일주소 from mdc_address where 부서 = 'BCC';"
        curs.execute(sSql)
        data = curs.fetchall()
        add_list = list(dict.fromkeys([i[0] for i in data]))
        mailtoBccList = ', '.join(add_list)
        return mailtoBccList

    def request_mails_summary(self):
        curs = self.conn.cursor()
        sSql = f"select mail, subject, sender, send_date, attachment, 처리현황, 처리일, 현장코드 " \
               f"   from mdc_mails " \
               f"   order by save_date desc " \
               f"   limit 300;"
        curs.execute(sSql)
        data = curs.fetchall()
        columns = ['mail', '제목', '보낸사람', '보낸날짜', '첨부파일', '처리현황', '처리일', '현장코드']
        df = pd.DataFrame(data, columns=columns)
        return df

    def request_site_code_by_site_name(self, site_name):
        curs = self.conn.cursor()
        sSql = f"select 현장코드, 현장명p from mdc_mst_site " \
               f"   where 현장명p = '{site_name}' " \
               f"   limit 1;"
        curs.execute(sSql)
        data = curs.fetchone()
        if not data:
            return None
        return data[0]

    def request_pension_result_data(self, site_code):
        curs = self.conn.cursor()
        sSql = f"select 분석월, 현장명p, 업체명, 출역일수, 소장출역, 직원출역, 확정일수 " \
               f"   from mdc_result " \
               f"   where 현장코드 = '{site_code}' ;"
        curs.execute(sSql)
        data = curs.fetchall()
        if len(data) == 0:
            return None
        columns = ['분석월', '현장명p', '업체명', '출역일수', '소장출역', '직원출역', '확정일수']
        df = pd.DataFrame(data, columns=columns)
        df['구분'] = df['업체명'].apply(lambda x: '당사' if '현대건설' in x else '협력업체')
        df['대비'] = df.apply(lambda x: str(round(x.확정일수 / x.출역일수 * 100))+"%" if x.출역일수 and x.확정일수 else "", axis=1)
        df['비고'] = df['확정일수'].apply(lambda x: "◎ 퇴직공제부금 미등록" if x == 0 else "")
        df = df.sort_values(['분석월', '구분', '업체명'], ascending=[False, True, True], kind='mergesort')
        return df[['분석월', '구분', '현장명p', '업체명', '출역일수', '소장출역', '직원출역', '확정일수', '대비', '비고']]

    def request_pension_result_data_full(self):
        curs = self.conn.cursor()
        sSql = f"select 현장코드, 분석월, 현장명p, 업체명, 출역일수, 소장출역, 직원출역, 확정일수 " \
               f"   from mdc_result ;"
        curs.execute(sSql)
        data = curs.fetchall()
        if len(data) == 0:
            return None

        columns = ['현장코드', '분석월', '현장명p', '업체명', '출역일수', '소장출역', '직원출역', '확정일수']
        df = pd.DataFrame(data, columns=columns)
        df['구분'] = df['업체명'].apply(lambda x: '당사' if '현대건설' in x else '협력업체')
        df['대비'] = df.apply(lambda x: round(x.확정일수 / x.출역일수, 2) if x.출역일수 and x.확정일수 else 0, axis=1)
        df['비고'] = df['확정일수'].apply(lambda x: "◎ 퇴직공제부금 미등록" if x == 0 else "")
        df = df.sort_values(['현장코드', '분석월', '구분', '업체명'], ascending=[True, False, True, True], kind='mergesort')
        df = df[['현장코드', '분석월', '구분', '현장명p', '업체명', '출역일수', '소장출역', '직원출역', '확정일수', '대비', '비고']]

        df = df.astype({
            "출역일수": int, "소장출역": int, "직원출역": int, "확정일수": float, "대비": float,
        })
        return df

    def request_site_code_and_site_name(self):
        curs = self.conn.cursor()
        sSql = f"SELECT distinct col_03 as 현장코드, col_04 as 현장명 " \
               f"   FROM prd_rawdata_md " \
               f"   where length(col_03) = 4 and length(col_04) > 3;"
        curs.execute(sSql)
        data = curs.fetchall()
        columns = ['현장코드', '현장명']
        df = pd.DataFrame(data, columns)
        return df

    def request_address_mdc_all(self):
        curs = self.conn.cursor()
        sSql = f"select 이름, 메일주소, 현장코드, 부서, 담당본부 " \
               f"   from mdc_address " \
               f"   order by 이름 ;"
        curs.execute(sSql)
        data = curs.fetchall()
        columns = ['이름', '메일주소', '현장코드', '부서', '담당본부']
        df = pd.DataFrame(data, columns=columns)
        return df


if __name__ == '__main__':

