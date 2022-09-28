import pandas as pd
import pytz

from model_data import *

class MailFormMaker:
    def __init__(self, dc, siteCode, month, bms_site_name=None):
        self.dc = dc
        self.siteCode = siteCode
        self.month = month
        if bms_site_name:
            self.bms_site_name = bms_site_name

    def write_subject(self):
        sSubject = f'작업일보 vs 전자카드 근로내역 출역정보 비교표 ({self.bms_site_name} / {self.month}월)'
        # sSubject = f'[퇴직공제부금/작업일보] {self.siteCode} / {self.bms_site_name}({sRunTime[0:13]}:{sRunTime[13:15]})'
        return sSubject

    def write_error_subject(self):
        sSubject = f'[작업일보 vs 전자카드 근로내역 출역정보 비교표] 등록된 현장명이 아니어서 처리되지 않았습니다. ({sRunTime[0:13]}:{sRunTime[13:15]})'
        return sSubject

    def consist_table_in_mail_body(self):
        """
        현장코드, 업체, 발송자, 날짜, 발송일시, 집계일시
        """
        dc = self.dc
        df = dc.request_table_in_mail_body(self.siteCode)
        htmlTable = """
                <table><thead>
                <tr style=\"height:30px;\"><td class=ment colspan=7 align=left>■ Daily Report Status</td></tr>
                <tr style=\"height:30px;\">
                <td class=head style=\"width:40px;\" align=center>No</td>
                <td class=head style=\"width:80px;\" align=center>Project</td>
                <td class=head style=\"width:300px;\" align=center>Subcontractor</td>
                <td class=head style=\"width:120px;\" align=center>Sender</td>
                <td class=head style=\"width:100px;\" align=center>Report Date</td>
                <td class=head style=\"width:100px;\" align=center>Received<br>[Local Time]</td>
                <td class=head style=\"width:100px;\" align=center>Summarized<br>[Korean Time]</td>
                </tr></thead>
                <tbody>
        """
        for index, row in df.iterrows():
            siteCode = row[0] # SM06
            subcon = row[1] # Daeah / Steel Structure
            mailSender = row[2] # Yoonjin Yoon <1300715@hdec.co.kr>
            mailSender = str(mailSender).replace('<', '[').replace('>', ']')
            reportDate = row[3] # 2022-03-03
            sendDate = row[4] # sendDate_attach = sendDate.strftime('%Y-%m-%d %H:%M:%S %z')
            try:
                sTimezone = dc.request_local_timezone(siteCode)
                sendDate = datetime.strptime(sendDate, '%Y-%m-%d %H:%M:%S %z')
                sendDate = sendDate.astimezone(pytz.timezone(sTimezone)).strftime('%Y-%m-%d %H:%M')
                # sendDate = datetime.strptime(sendDate, "%a, %d %b %Y %H:%M:%S %z").astimezone(
                #     pytz.timezone(sTimezone)).strftime('%Y-%m-%d %H:%M')
            except Exception as ex:
                logWrite(f"[타임 Formatter    ] : 시간대 변경 에러 // {ex}")
            SummarizedDate = row[5] # 2022-03-03_1415
            SummarizedDate = SummarizedDate[:10] + ' ' + SummarizedDate[11:13] + ':' + SummarizedDate[13:15]

            htmlTable += f"<tr style=\"height:30px;\">" \
                 f"<td class=data align=center>{index}</td>" \
                 f"<td class=data align=left>{siteCode}</td>" \
                 f"<td class=data align=left>{subcon}</td>" \
                 f"<td class=data align=left>{mailSender}</td>" \
                 f"<td class=data align=center>{reportDate}</td>" \
                 f"<td class=data align=center>{sendDate}</td>" \
                 f"<td class=data align=center>{SummarizedDate}</td>" \
                 f"</tr>"
        htmlTable += "</table><br><br><br>"

        return htmlTable


    def write_html_mail_body(self, htmlTable=None):
        # head
        sMailBody = """
        <html>
        <style>
        html {
        font-family:맑은 고딕, 돋움, 굴림;    
        color:black;
        }
        p {
        font-size:14px;
        font-weight:bold;
        }
        table, th, td {
        border: 1px solid black;
        border-collapse:collapse;
        height:25px;
        }
        table {
        font-size:12px;
        background:white;
        }
        td.head {
        font-weight:bold;
        color:white;
        background:#404040;
        }
        td.data {
        padding:10px;
        color:black;
        background:white;
        }
        td.ment {
        border: 1px solid white;
        font-size:13px;
        font-weight:bold;
        color:black;white;
        }
        </style>
        <body>
        <p>
        """
        # body
        sMailBody += f"본 메일은 당사 전자작업일보(HPMS, HCM, Easy작업일보)의 출역정보와 건설근로자공제회 전자카드근무관리시스템의 근로내역 비교표를 제공하기 위해 발송되는 메일입니다.<br><br>"
        sMailBody += f"첨부 파일 참조하시기 바랍니다.<br><br>"
        sMailBody += f"⊙ 현장코드 : {self.siteCode} <br>"
        sMailBody += f"⊙ 현 장 명 : {self.bms_site_name} <br>"
        # sMailBody += f"⊙ Date : {dateFrom} ~ {dateTo} <br><br>"
        # table
        if htmlTable is not None:
            sMailBody += htmlTable
        # tail
        sMailBody += "<br><br><br>"
        sMailBody += "※ Contact Point <br>"
        sMailBody += "&nbsp;&nbsp;&nbsp; 제도 관련 &nbsp;&nbsp;: 안전사업지원실 사업관리팀 (02-746-1940)<br>"
        sMailBody += "&nbsp;&nbsp;&nbsp; 데이터 관련 : 예산관리실 RM팀 (02-746-3643, 1339, 2262, 2782) <br>"
        sMailBody += """
        </p>
        </body>
        </html>
        """
        return sMailBody


if __name__ == '__main__':
    dc = DataControl()
    mailFormMaker = MailFormMaker(dc)
    sSubject = mailFormMaker.write_subject()
    htmlTable = mailFormMaker.consist_table_in_mail_body()
    sMailBody = mailFormMaker.write_html_mail_body(htmlTable)
    import os
    filename = 'mailbody.html'
    filename = os.path.join(savePath, filename)
    os.makedirs(savePath, exist_ok=True)
    with open(filename, 'w') as html_file:
        html_file.write(sMailBody)
    os.startfile(filename)