# -*- coding: utf-8 -*-

'''

IMAPを使用して他のメールクライアントでGmailをチェックする
https://support.google.com/mail/answer/7126229/?hl=ja

アプリパスワードでログイン
https://support.google.com/mail/answer/185833?hl=ja

2段階認証プロセスを有効にする
https://support.google.com/accounts/answer/185839?hl=ja

'''

from mail import Mail

class Gmail(Mail):

    def __init__(self, user, password):
        self.smtp_host = 'smtp.gmail.com'
        self.smtp_port = 465
        self.smtp_auth = True
        self.smtp_ssl  = True
        self.smtp_tls  = False
        self.smtp_from = '%s@gmail.com'
        self.imap_host = 'imap.gmail.com'
        self.imap_port = 993
        self.imap_ssl  = True
        self.imap_tls  = False
        Mail.__init__(self, user, password)
