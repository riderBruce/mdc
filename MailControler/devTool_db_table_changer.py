from model_data import *


def insert_site_info():
    dc = DataControl()
    curs = dc.conn.cursor()
    data = [
        ['IBRU', 'IRAQ BASRA REFINERY UPGRADING PROJECT', 'BASRA', 'IRQ', 'Asia/Baghdad'],

            ]
    table = "job_address"
    for i in data:
        sSql = f"insert into {table} " \
               f"(현장코드, 현장명, 현장별칭, 국가, 타임존) " \
               f"values(%s,%s,%s,%s,%s)"
        try:
            curs.execute(sSql, i)
        except Exception as ex:
            print(ex, sSql)

def insert_mail_address():
    dc = DataControl()
    curs = dc.conn.cursor()
    data = [
        ['SM12', 'SAUDI MARJAN INCREMENT PROGRAM PKG 12', '1300715', '윤윤진', '', 'guillaume.yy@hdec.co.kr', '+821091418486'],
    ]
    table = "job_address"
    for i in data:
        sSql = f"insert into {table} " \
               f"(현장코드, 현장명, 사번, 이름, 직책, 메일주소, 휴대전화) " \
               f"values(%s,%s,%s,%s,%s,%s,%s)"
        try:
            curs.execute(sSql, i)
        except Exception as ex:
            print(ex, sSql)

def insert_mail_address_cc():
    dc = DataControl()
    curs = dc.conn.cursor()
    data = [
        ['SM12', 'SAUDI MARJAN INCREMENT PROGRAM PKG 12', '1300715', '윤윤진', '', 'guillaume.yy@hdec.co.kr', '+821091418486'],
    ]
    table = "job_address_cc"
    for i in data:
        sSql = f"insert into {table} " \
               f"(현장코드, 현장명, 사번, 이름, 직책, 메일주소, 휴대전화) " \
               f"values(%s,%s,%s,%s,%s,%s,%s)"
        try:
            curs.execute(sSql, i)
        except Exception as ex:
            print(ex, sSql)


if __name__ == '__main__':
    insert_mail_address()
    insert_mail_address_cc()
    insert_site_info()
    dc.conn.commit()