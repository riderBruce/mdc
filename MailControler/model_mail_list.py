
class MailList:
    def __init__(self, dc, MAIL, siteCode):
        self.MAIL = MAIL
        self.siteCode = siteCode

        self.dc = dc

    def select_mailing_list(self):
        dc = self.dc

        if self.MAIL == 'ALL':
            mailtoList = dc.request_address_mailto(self.siteCode)
            mailtoccList = dc.request_address_cc()
            mailtoBccList = dc.request_address_bcc()
        else:
            mailtoList = self.MAIL
            mailtoccList = None
            mailtoBccList = None

        return mailtoList, mailtoccList, mailtoBccList