# -*- coding: utf-8 -*-

from __future__ import unicode_literals

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

from common import log

class MailHandler:

    def __init__(self, user, password):
        self.user = user
        self.password = password
        self.email_default_encoding = 'utf-8'

    def get_from_address(self):
        return self.smtp_from % self.user

    def send(self, subject, body, to_addresses, cc_addresses=None, bcc_addresses=None, replyto_address=None):
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
        from_address = self.smtp_from % self.user
        msg = MIMEText(body, 'plain', self.email_default_encoding)
        msg['Subject'] = Header(subject, self.email_default_encoding)
        msg['From'] = from_address
        msg['To'] = ', '.join(to_addresses)
        if cc_addresses and len(cc_addresses) > 0:
            msg['CC'] = ', '.join(cc_addresses)
        if bcc_addresses and len(bcc_addresses) > 0:
            msg['BCC'] = ', '.join(bcc_addresses)
        if replyto_address:
            msg['Reply-To'] = replyto_address;
        msg['Date'] = Utils.formatdate(localtime=True)
        conn.sendmail(from_address, to_addresses, msg.as_string())
        conn.close()

    def receive(self, criterion='ALL', pref='TEXT'):
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
        lastid = None
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
            # append to array
            emails.append(item)
            # keep id as lastid
            lastid = id
        conn.close()
        conn.logout()
        # sort array in reverse order of Date
        return lastid, sorted(emails, key=lambda item: email.utils.mktime_tz(email.utils.parsedate_tz(item['Date'])), reverse=True)

    def convert_header(self, msg, key):
        try:
            buf = []
            value = msg.get(key).replace('\r\n','').replace('==?=','==?= ')
            if value:
                for segment in decode_header(value):
                    text = segment[0]
                    encoding = segment[1]
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
