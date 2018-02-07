# -*- coding: utf-8 -*-

'''

iCloud メールクライアント向けのメールサーバ設定
https://support.apple.com/ja-jp/HT202304

App 用パスワードを使う
https://support.apple.com/ja-jp/HT204397

'''

from mailhandler import MailHandler

class iCloud(MailHandler):

    def __init__(self, user, password):
        self.smtp_host = 'smtp.mail.me.com'
        self.smtp_port = 587
        self.smtp_auth = True
        self.smtp_ssl  = False
        self.smtp_tls  = True
        self.smtp_from = '%s@icloud.com'
        self.imap_host = 'imap.mail.me.com'
        self.imap_port = 993
        self.imap_ssl  = True
        self.imap_tls  = False
        MailHandler.__init__(self, user, password)
