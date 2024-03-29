import poplib
import email
import os
import re
from datetime import datetime, timedelta

from model_data import *

class EmailAttachDownloader:

    def __init__(self, dc):
        self.pop3_server =
        self.pop3_user =
        self.pop3_passwd =
        self.p = poplib.POP3(self.pop3_server)
        self.dc = dc

    def check_emails(self):
        # # 상위폴더까지 생성 / 폴더 존재시 에러 없음
        # os.makedirs(savePath, exist_ok=True)
        # pop3 server check
        self.p.user(self.pop3_user) # b'+OK'
        self.p.pass_(self.pop3_passwd) # b'+OK User successfully logged on.'
        mailCount, mailBoxSize = self.p.stat()  # (851, 573002870)
        logWrite(f'[메일박스 - Check   ] 메일 개수 : {mailCount} ea  / 메일 용량 : {mailBoxSize} bytes')
        # get all message number
        msg_list = self.p.list()  # (b'+OK 851 573002870', [b'1 862025', ~~  (헤더, 본문, ~) 3개로 구성됨
        if not msg_list[0].decode("utf-8").startswith('+OK'):
            logWrite("★[서버오류 - Check   ] error !!!!")
            exit(1)
        email_num = [int(msg.split()[0].decode("utf-8")) for msg in msg_list[1]]
        return email_num

    def download_attachFiles(self, email_num):
        # 최근 X개 메일, 그중에 X일 미만된 메일만 실행
        before_mail = 300
        if len(email_num) == 0:
            return []
        from_mail_num = email_num[-1] - before_mail
        from_mail_num = from_mail_num if from_mail_num > 0 else 0
        email_num = email_num[from_mail_num:]
        before_day = 100

        mail_info = []
        for n in email_num:
            # read each mail ---------------------------------------------------------
            response = self.p.retr(n)
            # response = p.retr(800)
            # p.retr(852) = (b'+OK message follows', [b'Delivered-To: pahkey@gmail.com', ... 생략 ...,  b''], 4029)
            # -------------------------------------------------------------------
            # decoding code check
            responseStr = ""
            try:
                responseStr = [data.decode('utf-8') for data in response[1]]
                is_encoding = True
            except:
                is_encoding = False
            if is_encoding == False:
                try:
                    responseStr = [data.decode('euc-kr') for data in response[1]]
                    is_encoding = True
                except:
                    is_encoding = False
            if is_encoding == False:
                logWrite(f'★[메일읽기 - fail    ] 인코딩 실패 : {str(response[1][:40])}')
                continue

            # # decode -------------------------------------------------------------
            parsed_msg = email.message_from_string('\n'.join(responseStr))

            # Date
            dateWithCode = email.header.decode_header(parsed_msg['Date'])
            code = dateWithCode[0][1]  # define Encoding Code if it is
            if code:
                sendDate = dateWithCode[0][0].decode(code, errors='ignore')
            else:
                sendDate = dateWithCode[0][0]
            # select new mail by date
            sendDate = datetime.strptime(str(sendDate[:31]).strip(), "%a, %d %b %Y %H:%M:%S %z")
            # sendDate_mail = sendDate.strftime("%a, %d %b %Y %H:%M:%S %z")
            # sendDate_attach = sendDate.strftime('%Y-%m-%d %H:%M:%S %z')
            if datetime.now().date() - sendDate.date() > timedelta(days=before_day):
                print(f'[메일읽기 - Skip    ] : No.{n} - 1주일 이상된 메일입니다.')
                continue

            # Subject
            subjectWithCode = email.header.decode_header(parsed_msg['Subject'])
            code = subjectWithCode[0][1] # define Encoding Code if it is
            if code:
                subject = subjectWithCode[0][0].decode(code, errors='ignore')
            else:
                subject = subjectWithCode[0][0]

            # From - list structure that need for loop
            senderWithCode = email.header.decode_header(parsed_msg['From'])
            senderTxt = ""
            for rowSender in senderWithCode:
                code = rowSender[1] # define Encoding Code if it is
                if code:
                    senderTxt = senderTxt + rowSender[0].decode(code, errors='ignore')
                else:
                    senderTxt = senderTxt + str(rowSender[0])

            # mailStr
            subject = subject.replace('\'', '')
            # sendDate = sendDate.replace('\'', '')
            senderTxt = senderTxt.replace('\'', '').replace('b <', ' <')
            # mailStr = subject + '|' + senderTxt + '|' + str(sendDate)
            sendDate = sendDate.strftime('%Y-%m-%d %H:%M:%S %z')
            mailStr = f"{str(sendDate)[:19]} | {senderTxt[:3]} | {subject[:10]}"

            # -------------------------------------------------------------------
            # new mail by check DB

            if self.dc.is_new_mail(mailStr):
                logWrite(f'[메일읽기 - New     ] : {mailStr}')
                save_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S') + " +0900"
                attachment = False
                for part in parsed_msg.walk():

                    # logWrite(part.get_content_maintype())
                    # 첨부파일 중 message는 content-disposition 항목이 있지만, 첨부파일 제목은 없음
                    if part.get_content_maintype() in ['multipart', 'message']:
                        continue
                    if not part.get('Content-Disposition'):
                        continue
                    # ex)
                    # Content - Disposition: attachment;
                    # filename = "=?UTF-8?B?MjIuMDEuMTRf7J6R7JeF7J2867O0X+2IrOyciOyKpOy7tF/ssqjri6gueGxzeA==?="
                    # logWrite(f"{part.get_filename()}, {type(part.get_filename())}, {n}")
                    filenameDecode = email.header.decode_header(part.get_filename())
                    code = filenameDecode[0][1]
                    if code:
                        filename = filenameDecode[0][0].decode(code)
                    else:
                        filename = filenameDecode[0][0]
                    filename = re.sub('[^\w\s-]','.', filename).strip().lower()
                    filename = filename.replace('\n', '').replace('\r', '').replace('\t', '')
                    filename = os.path.join(savePath, filename)
                    # while os.path.exists(filename) == True:
                    #     filename_tmp, file_extension = os.path.splitext(filename)
                    #     filename = filename_tmp + '_' + datetime.now().strftime('%Y%m%d%H%M%S') + file_extension

                    # confirm the file is already finished
                    filename_stem = Path(filename).stem
                    if self.dc.is_already_finished(mailStr, sendDate, senderTxt, subject, save_date, filename_stem):
                        logWrite(f"[▷ 이미 종료된 파일명   ] : {mailStr}, {filename_stem}, 처리하지 않습니다. ")
                        continue

                    # save attachment
                    fp = open(filename, 'wb')
                    fp.write(part.get_payload(decode=1))
                    fp.close()
                    logWrite('[첨부파일 - Save    ] : ' + filename)

                    # insert mail list into DB
                    self.dc.insert_mail_get_list(mailStr, sendDate, senderTxt, subject, save_date, filename_stem)
                    logWrite(f'[메일로그 - Save    ] : {mailStr} /// {sendDate}  /// {senderTxt} ///  {subject} /// {filename_stem} /// {save_date}')
                    attachment = True
                    # self.dc.update_attachment(mailStr, filename, save_date)
                    # logWrite(f'[첨부로그 - Save    ] : {mailStr} /// {filename} /// {save_date}')
                    mail_info.append([filename, sendDate, senderTxt, subject, save_date])

                if attachment == False:
                    self.dc.insert_mail_get_list(mailStr, sendDate, senderTxt, subject, save_date)
                    self.dc.update_mails_process_status_by_mailStr(mailStr, sendDate, senderTxt, save_date, "첨부파일없음")
                    logWrite(f'[메일로그 - Save    ] : {mailStr} /// {sendDate}  /// {senderTxt} ///  {subject} /// {save_date}')

            else:
                print('[메일읽기 - Skip    ] : ' + mailStr)

        self.p.quit()

        return mail_info


if __name__ == '__main__':
    dc = DataControl()
    jobReportDownloader = EmailAttachDownloader(dc)
    email_num = jobReportDownloader.check_emails()
    mail_info = jobReportDownloader.download_attachFiles(email_num)
