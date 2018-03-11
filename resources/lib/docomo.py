# -*- coding: utf-8 -*-

'''

他社製品の携帯電話機、または他のメールソフトからのご利用方法
https://www.nttdocomo.co.jp/service/docomo_mail/other/#p02

IMAP専用パスワードの設定方法
https://www.nttdocomo.co.jp/service/docomo_mail/other/security/#p04

'''

from mail import Mail

class Docomo(Mail):

    def __init__(self, user, password, smtp_from):
        self.smtp_host = 'smtp.spmode.ne.jp'
        self.smtp_port = 465
        self.smtp_auth = True
        self.smtp_ssl  = True
        self.smtp_tls  = False
        self.smtp_from = smtp_from
        self.imap_host = 'imap.spmode.ne.jp'
        self.imap_port = 993
        self.imap_ssl  = True
        self.imap_tls  = False
        Mail.__init__(self, user, password)
