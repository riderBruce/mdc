import os
import smtplib
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders

from model_data import *

class MailSender:
    def __init__(self, dc, siteCode=None):
        self.dc = dc
        if siteCode:
            self.siteCode = siteCode
        else:
            self.siteCode = "----"
        self.smtp_server =  # 메일서버
        self.smtp_user =  # ID
        self.smtp_passwd =   # 비밀번호
        self.smtp_sender =   # 이메일

    def send_mail_smtp(self, sSubject, sMailBody, mailtoList, mailtoccList, mailtoBccList, outputFileName=None):
        try:
            server = smtplib.SMTP(self.smtp_server)
            server.ehlo()
            server.starttls()
            server.login(self.smtp_user, self.smtp_passwd)
        except:
            logWrite('★[메일서버 접속오류] : SMTP서버 ')
        # write message
        msg = MIMEBase("multipart", "mixed")
        msg["Subject"] = sSubject
        msg["From"] = self.smtp_sender
        if mailtoList is not None:
            msg["To"] = mailtoList
        if mailtoccList is not None:
            msg["Cc"] = mailtoccList
        if mailtoBccList is not None:
            msg["Bcc"] = mailtoBccList
        msg.attach(MIMEText(sMailBody, 'html', _charset='utf-8'))
        if outputFileName:
            # attach File
            # part = MIMEBase('application', 'octet-stream')
            part = MIMEBase('application', 'vnd.ms-excel')
            part.set_payload(open(outputFileName, 'rb').read())
            file_name = os.path.basename(outputFileName)

            # # xls: 파일 제목만, xlsx: 전체
            # filename_tmp, file_extension = os.path.splitext(file_name)
            # if file_extension == ".xls":
            #     file_name = Path(file_name).stem

            part.add_header('Content-Description', file_name)
            part.add_header('Content-Disposition', 'attachment', filename=file_name)
            encoders.encode_base64(part)
            msg.attach(part)
        # send mail
        try:
            server.send_message(msg)
            logWrite(f"[메일발송 - Send    ] : {self.siteCode} / {sSubject} / mailto: {mailtoList} / cc: {mailtoccList} / bcc: {mailtoBccList}")
        except Exception as ex:
            logWrite(f"[메일발송 - Error   ] {self.siteCode} / {sSubject} / mailto: {mailtoList} / cc: {mailtoccList} / bcc: {mailtoBccList}")
            logWrite(f"[메일발송 - Error   ] {str(ex)}")
        server.quit()



if __name__ == '__main__':
    pass
