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
            siteCode = row[0] # 
            subcon = row[1] # 
            mailSender = row[2] # 
            mailSender = str(mailSender).replace('<', '[').replace('>', ']')
            reportDate = row[3] # 
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

        </p>
        </body>
        </html>
        """
        return sMailBody


if __name__ == '__main__':
