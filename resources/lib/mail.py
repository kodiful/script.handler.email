# -*- coding: utf-8 -*-

import re
import chardet

import email
import email.utils
from email.header import decode_header
from email.Header import Header
from email.MIMEText import MIMEText
from email import Utils
from imaplib import IMAP4,IMAP4_SSL
from smtplib import SMTP,SMTP_SSL

class Mail:

    def __init__(self, user, password, config=None):
        self.user = user
        self.password = password
        if config:
            self.smtp_host = config['smtp_host']
            self.smtp_port = config['smtp_port']
            self.smtp_auth = config['smtp_auth']
            self.smtp_ssl  = config['smtp_ssl']
            self.smtp_tls  = config['smtp_tls']
            self.smtp_from = config['smtp_from']
            self.imap_host = config['imap_host']
            self.imap_port = config['imap_port']
            self.imap_ssl  = config['imap_ssl']
            self.imap_tls  = config['imap_tls']
        if self.smtp_from.find('%s') > -1:
            self.smtp_from = self.smtp_from % user
        self.email_default_encoding = 'utf-8'

    def send(self, subject, body, to, cc=[], bcc=[], replyto=None):
        if self.smtp_ssl:
            conn = SMTP_SSL(self.smtp_host, self.smtp_port)
        else:
            conn = SMTP(self.smtp_host, self.smtp_port)
        if self.smtp_tls:
            conn.ehlo()
            conn.starttls()
            conn.ehlo()
        if self.smtp_auth:
            conn.login(self.user, self.password)
        msg = MIMEText(body, 'plain', self.email_default_encoding)
        msg['Subject'] = Header(subject, self.email_default_encoding)
        msg['From'] = self.smtp_from
        msg['To'] = ', '.join(to)
        if len(cc) > 0:
            msg['CC'] = ', '.join(cc)
        if replyto:
            msg['Reply-To'] = replyto
        msg['Date'] = Utils.formatdate(localtime=True)
        conn.sendmail(self.smtp_from, to+cc+bcc, msg.as_string())
        conn.close()

    def receive(self, criterion='ALL', pref='TEXT'):
        if isinstance(criterion, str): criterion = criterion.decode('utf-8')
        emails = []
        if self.imap_ssl:
            conn = IMAP4_SSL(self.imap_host, self.imap_port)
        else:
            conn = IMAP4(self.imap_host, self.imap_port)
        if self.imap_tls:
            conn.ehlo()
            conn.starttls()
            conn.ehlo()
        conn.login(self.user, self.password)
        conn.select('inbox')
        typ, msgs = conn.search(None, criterion)
        for id in msgs[0].split():
            typ, data = conn.fetch(id, '(RFC822)')
            msg = email.message_from_string(data[0][1])
            # parse message body
            if msg.is_multipart():
                text = None
                html = None
                for msg1 in msg.get_payload():
                    if msg1.get_content_type() == 'text/plain':
                        text = self.convert_payload(msg1)
                    if msg1.get_content_type() == 'text/html':
                        html = self.convert_payload(msg1)
                    if msg1.get_content_type() == 'multipart/alternative':
                        for msg2 in msg1.get_payload():
                            if msg2.get_content_type() == 'text/plain':
                                text = self.convert_payload(msg2)
                            if msg2.get_content_type() == 'text/html':
                                html = self.convert_payload(msg2)
                if pref == 'TEXT':
                    if not text is None:
                        body = text.strip()
                    elif not html is None:
                        body = html.strip()
                    else:
                        body = '(null)'
                if pref == 'HTML':
                    if not html is None:
                        body = html.strip()
                    elif not text is None:
                        body = text.strip()
                    else:
                        body = '(null)'
            else:
                body = self.convert_payload(msg)
            # pack data
            item = {'body':body}
            for key in ['From','To','CC','Reply-To','Subject']:
                item[key] = self.convert_header(msg,key)
            for key in ['Date','Message-Id']:
                item[key] = msg.get(key)
            # normalize Message-Id
            item['Message-Id'] = re.sub(r'[^a-zA-Z0-9\-\.\@\_]', '', item['Message-Id'])
            # append to array
            emails.append(item)
        conn.close()
        conn.logout()
        # sort array in reverse order of Date
        return sorted(emails, key=lambda item: email.utils.mktime_tz(email.utils.parsedate_tz(item['Date'])), reverse=True)

    def convert_header(self, msg, key):
        try:
            buf = []
            value = msg.get(key)
            segments = value.replace('?=','?=\r\n').split('\r\n')
            for s1 in segments:
                for s2 in decode_header(s1):
                    text = s2[0]
                    encoding = s2[1]
                    if not encoding is None:
                        text = text.decode(encoding,'replace')
                    buf.append(text)
            return ' '.join(buf).strip()
        except:
            return ''

    def convert_payload(self, msg):
        try:
            charset = msg.get_content_charset()
            if charset is None:
                charset = chardet.detect(str(msg))['encoding']
            return unicode(msg.get_payload(decode=True),str(charset),'replace').strip()
        except:
            return ''
