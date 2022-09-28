
import psycopg2 as pg2

db_host = '10.171.94.66'
db_name = 'JobReport'
db_user = 'postgres'
db_pwd = 'nam1004'
conn = pg2.connect(
    'host={0} dbname={1} user={2} password={3}'.format(db_host, db_name, db_user,
                                                       db_pwd))
curs = conn.cursor()
sSql = """
update mdc_mails set 
처리현황 = '담당자확인/처리완료' , 처리일 = '2022-08-03 10:20:00' 
where (send_date = '2022-08-02 16:46:01 +0900' or send_date = '2022-08-02 16:49:25 +0900') 
and 현장코드 is not null 
and  처리현황 is null;
"""
curs.execute(sSql)
conn.commit()
print(curs.rowcount, "rows effected")
