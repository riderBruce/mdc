import sys
import time
from tqdm.auto import tqdm

sys.path.append(r"D:\Project_Python\webMDChecker\MDChecker\MailControler")

from model_mail_get import EmailAttachDownloader
from model_excel_export import SummaryExcelData
from model_excel_import import ExcelDataConverter
from model_data import *
from model_data_converter import PensionDataConverter
from model_excel_styler import ExcelFormMaker
from model_mail_styler import MailFormMaker
from model_mail_list import MailList
from model_mail_send import MailSender
from conf import config

# ----------------------------------------
logWrite('────────────────────── ★ Run  Start ★ ───────────────────────')

sArgv = ""     # DEV, DEV_LOCAL, ADMIN, SERVER conf.py 파일에서 분기를 잡아준다.

if len(sys.argv) > 1:
    sArgv = sys.argv[1]
else:
    sArgv = "DEV_LOCAL"
# sArgv = "SERVER"

DB = config[sArgv]['DB']
MAIL = config[sArgv]['MAIL']

try:
    logWrite(f"sArgv: {sArgv}, DB: {DB}, MAIL: {MAIL}")
except Exception as e:
    logWrite(e)


# ----------------------------------------
startTime = datetime.now()
logWrite(f'start : {startTime}')
# close all opened excel
kill_excel()
# db controller
dc = DataControl(DB)

# dc.delete_old_data('mdc_mails')
# dc.delete_old_data('mdc_raw_md')
# dc.delete_old_data('mdc_result')
# dc.delete_old_data('mdc_mails', "where mail = '2022-09-27 08:18:54 | 최다희 | 7월 추가 현장 테';")
# dc.delete_old_data('mdc_mails', "where mail = '2022-08-02 16:46:01 | 최다희 | 4월 전체 송부 (';")
# dc.delete_old_data('mdc_mails', "where mail = '2022-07-06 | 김영일 | 5월 자료';")

# # email check
ed = EmailAttachDownloader(dc)
email_num = ed.check_emails()
# delete old attachment / attached file download / write result in DB
mail_info = ed.download_attachFiles(email_num)

if len(mail_info) == 0:
    logWrite(f'[신규 메일 없음      ] : 신규 메일이 없으므로 프로세스를 종료됩니다. ')

for fileName, send_date, senderTxt, subject, save_date in tqdm(mail_info, position=1, leave=True):
    fileName_stem = Path(fileName).stem

    # file data save into db
    ec = ExcelDataConverter(fileName)
    if not ec.is_excel_file():
        dc.update_mails_add_site_code_None(fileName_stem, send_date)
        dc.update_mails_process_status(fileName_stem, send_date, senderTxt, "첨부파일오류")
        dc.conn.commit()
        continue
    sheetName, usedRangeData = ec.read_excel_usedRangeData(1)
    if not ec.is_pension_file(usedRangeData):
        dc.update_mails_add_site_code_None(fileName_stem, send_date)
        dc.update_mails_process_status(fileName_stem, send_date, senderTxt, "첨부파일오류")
        dc.conn.commit()
        continue
    df, month = ec.convert_pension_data_for_DB(sheetName, send_date, usedRangeData)
    if df.empty:
        dc.update_mails_add_site_code_None(fileName_stem, send_date)
        dc.update_mails_process_status(fileName_stem, send_date, senderTxt, "데이터없음")
        dc.conn.commit()
        continue
    dc.drop_duplicates_by_date_from_raw(df['공제가입번호'].unique(), month)
    dc.insert_data_to_db(df, 'mdc_raw_md')

    # update mdc_mails
    dc.update_mails_sheetName(fileName_stem, sheetName, send_date)
    logWrite(f'[데이터 저장 - Save ] : {fileName_stem} ')

    # data convert by fileName
    pc = PensionDataConverter(dc, fileName_stem, month)
    siteCode, df_converted = pc.pension_data_converter()

    # update_mails_table
    if siteCode is not None:
        # 분석데이터 입력 / 기존 중복 데이터 삭제 (월, 현장코드)
        dc.delete_old_data('mdc_result', f"where 분석월 = '{month}' and 현장코드 = '{siteCode}'")
        dc.insert_data_to_db(df_converted, 'mdc_result')
        # 하나의 파일에 있는 신규 현장명 자동 등록 : mdc_mst_site : site_code / site_name_p upload
        dc.insert_new_site_name_to_mdc_mst_site(df_converted)
        # mails에 현장코드 자동 등록
        dc.update_mails_add_site_code(fileName_stem, send_date, senderTxt)
        # 메일 주소 자동 등록
        dc.insert_new_mail_sender_address(siteCode, senderTxt)
        # 이메일 중복 제거
        columns = ['이름', '메일주소', '현장코드', '부서', '담당본부']
        dc.drop_duplicates_from_DB_table('mdc_address', columns)

        logWrite(f'[데이터 가공 - Save ] : {fileName_stem} ')

        # reply : excel / styling / send mail
        se = SummaryExcelData(dc, siteCode, month)
        outputFileName = se.name_output_excel_file()
        se.excel_writer(outputFileName)

        # form excel : merge, color, align – win32com.client
        bms_site_name = dc.request_site_name_by_site_code(siteCode)
        ExcelFormMaker(outputFileName, siteCode, bms_site_name, month)

        # make mailBody
        mf = MailFormMaker(dc, siteCode, month, bms_site_name)
        sSubject = mf.write_subject()
        # htmlTable = mf.consist_table_in_mail_body()
        # sMailBody = mf.write_html_mail_body(htmlTable)
        sMailBody = mf.write_html_mail_body()

        # make mailinglist
        ml = MailList(dc, MAIL, siteCode)
        mailtoList, mailtoccList, mailtoBccList = ml.select_mailing_list()

        # send mail - smtp
        mailSender = MailSender(dc, siteCode)
        mailSender.send_mail_smtp(sSubject, sMailBody, mailtoList, mailtoccList, mailtoBccList, outputFileName)

        # update status
        dc.update_mails_older_process_status_by_siteCode(fileName_stem, send_date, senderTxt, siteCode)
        dc.update_mails_process_status_by_siteCode(siteCode, save_date, "메일송부완료")

        # 파일 1개 처리하여 메일 송부 후
        dc.conn.commit()

    else:
        # site_code = None
        dc.update_mails_add_site_code_None(fileName_stem, send_date)
        logWrite(f"{fileName_stem}_{send_date} : 등록되지 않은 현장 자료입니다.")

        # make mailBody
        bms_site_name = dc.request_site_name_by_site_code(siteCode)
        mf = MailFormMaker(dc, siteCode, month, bms_site_name)
        sSubject = mf.write_error_subject()
        un_matching_list = df.현장명p.to_list()
        bodyText = f"{send_date} : 해당 일시에 수신된 데이터를 매칭할 수 없습니다. 확인하시기 바랍니다. <br><br>" \
                   f"{un_matching_list}"
        sMailBody = mf.write_html_mail_body(bodyText)

        # make mailinglist
        ml = MailList(dc, MAIL, siteCode)
        mailtoList, mailtoccList, mailtoBccList = ml.select_mailing_list()

        # send mail - smtp
        mailSender = MailSender(dc, siteCode)
        mailSender.send_mail_smtp(sSubject, sMailBody, mailtoList, mailtoccList, mailtoBccList, fileName)
        dc.update_mails_process_status_to_ongoing(fileName_stem, send_date, "담당자확인중")

        # 파일 1개 처리하여 메일 송부 후
        dc.conn.commit()

try:
    time.sleep(1)
    shutil.rmtree(savePath)
    logWrite(f'[첨부파일 - 전체삭제  ] : {savePath}')
except OSError as e:
    logWrite(f'[첨부파일 - 삭제오류  ] : {e.filename}, {e.strerror}')

logWrite(f'finish : {datetime.now() - startTime}')
logWrite('────────────────────── ★ Run Finish ★ ───────────────────────')



