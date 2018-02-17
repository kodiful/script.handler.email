# -*- coding: utf-8 -*-

import xbmc
import xbmcgui
import xbmcplugin
import xbmcaddon

import os
import urllib, urlparse
import re
import email.utils
import time, datetime

from bs4 import BeautifulSoup

from resources.lib.common import log, notify, formatted_datetime

from resources.lib.mail import Mail
from resources.lib.gmail import Gmail
from resources.lib.icloud import iCloud

class Main:

    def __init__(self, service=None):
        # アドオン
        self.addon = xbmcaddon.Addon()
        # メールサービスを初期化
        self.service = None
        service = service or self.addon.getSetting('service')
        if service == 'Custom':
            user      = self.addon.getSetting('user')
            password  = self.addon.getSetting('password')
            smtp_host = self.addon.getSetting('smtp_host')
            smtp_port = self.addon.getSetting('smtp_port')
            smtp_auth = self.addon.getSetting('smtp_auth')
            smtp_ssl  = self.addon.getSetting('smtp_ssl')
            smtp_tls  = self.addon.getSetting('smtp_tls')
            smtp_from = self.addon.getSetting('smtp_from')
            imap_host = self.addon.getSetting('imap_host')
            imap_port = self.addon.getSetting('imap_port')
            imap_ssl  = self.addon.getSetting('imap_ssl')
            imap_tls  = self.addon.getSetting('imap_tls')
            if user and password and smtp_host and smtp_port and smtp_from and smtp_host and smtp_port:
                config = {
                    'smtp_host': smtp_host,
                    'smtp_port': int(smtp_port),
                    'smtp_auth': smtp_auth=='true',
                    'smtp_ssl':  smtp_ssl=='true',
                    'smtp_tls':  smtp_tls=='true',
                    'smtp_from': smtp_from,
                    'imap_host': imap_host,
                    'imap_port': int(imap_port),
                    'imap_ssl':  imap_ssl=='true',
                    'imap_tls':  imap_tls=='true'
                }
                self.service = Mail(user, password, config)
        elif service == 'Gmail':
            user = self.addon.getSetting('user1')
            password = self.addon.getSetting('password1')
            if user and password:
                self.service = Gmail(user, password)
        elif service == 'iCloud':
            user = self.addon.getSetting('user2')
            password = self.addon.getSetting('password2')
            if user and password:
                self.service = iCloud(user, password)
        # メールサービスが正常に初期化されたら他の初期化を実行
        if self.service:
            # キャッシュディレクトリのパス
            profile_path = xbmc.translatePath(self.addon.getAddonInfo('profile'))
            self.cache_path = os.path.join(profile_path, 'cache', service)
            if not os.path.isdir(self.cache_path):
                os.makedirs(self.cache_path)
            # 表示するメール数
            self.listsize = self.addon.getSetting('listsize')
            if self.listsize == 'Unlimited':
                self.listsize = 0
            else:
                self.listsize = int(self.listsize)

    def main(self):
        # パラメータ抽出
        params = {'action':'','filename':'','subject':'','message':'', 'to':[], 'cc':[]}
        args = urlparse.parse_qs(sys.argv[2][1:])
        for key in params.keys():
            params[key] = args.get(key, params[key])
        for key in ['action','filename','subject','message']:
            params[key] = params[key] and params[key][0]
        # メイン処理
        if params['action'] == '':
            start = self.addon.getSetting('start')
            if start == "true":
                # 新着をチェックして表示
                self.check()
            else:
                # 新着をチェックしないで表示
                self.list(newmails=[])
        elif params['action'] == 'check':
            # 新着をチェックして表示
            self.check()
        elif params['action'] == 'refresh':
            # キャッシュクリア
            files = os.listdir(self.cache_path)
            for filename in files:
                os.remove(os.path.join(self.cache_path, filename))
            # 再読み込み
            xbmc.executebuiltin('Container.Update(%s?action=check,replace)' % (sys.argv[0]))
        elif params['action'] == 'open':
            # メールの内容を表示
            if params['filename']: self.open(params['filename'])
        elif params['action'] == 'sendmessage':
            # メールを送信
            subject = self.addon.getSetting('subject')
            message = self.addon.getSetting('message')
            to = self.addon.getSetting('to')
            cc = self.addon.getSetting('cc')
            # 送信データ
            values = {'action':'send', 'subject':subject, 'message':message, 'to':to, 'cc':cc}
            postdata = urllib.urlencode(values)
            xbmc.executebuiltin('RunPlugin(%s?%s)' % (sys.argv[0], postdata))
        elif params['action'] == 'send':
            # メールを送信
            if self.addon.getSetting('bcc') == 'true':
                bcc = [self.service.smtp_from]
            else:
                bcc = None
            self.send(subject=params['subject'], message=params['message'], to=params['to'], cc=params['cc'], bcc=bcc)

    def check(self, refresh=True):
        # 管理用ファイル
        criterion_file = os.path.join(self.cache_path, '.criterion')
        newmails_file = os.path.join(self.cache_path, '.newmails')
        # 前回表示時の月日を読み込む
        if os.path.isfile(criterion_file):
            f = open(criterion_file, 'r')
            criterion = f.read()
            f.close()
        else:
            # 前回表示時の月日が不明の場合は30日前に設定
            d = datetime.datetime.utcnow() - datetime.timedelta(days=30)
            criterion = d.strftime('SINCE %d-%b-%Y')
        # 設定した期間のメールを検索
        newmails, senders = self.receive(criterion)
        if len(newmails) > 0:
            # 新着メールのファイルパスリストを書き出す
            f = open(newmails_file, 'a')
            f.write('\n'.join(newmails) + '\n')
            f.close()
            xbmc.sleep(1000)
            # Kodiをアクティベート
            if self.addon.getSetting('cec') == 'true':
                xbmc.executebuiltin('XBMC.CECActivateSource')
            # 新着があることを通知
            notify('New mail from %s' % senders[newmails[0]])
        # アドオン操作で呼び出された場合の処理
        if refresh:
            if os.path.isfile(newmails_file):
                # 新着メールのファイルパスリストを読み込む
                f = open(newmails_file, 'r')
                newmails = f.read().split('\n')
                f.close()
                # 新着メールのファイルパスリストを削除
                os.remove(newmails_file)
            else:
                newmails = []
            # リストを表示
            self.list(newmails)
            # 次回表示時のために今回表示時の月日を書き出す
            d = datetime.datetime.utcnow()
            criterion = d.strftime('SINCE %d-%b-%Y')
            f = open(criterion_file, 'w')
            f.write(criterion)
            f.close()

    def send(self, subject, message, to, cc=[], bcc=[]):
        # メール送信
        self.service.send(subject, message, to, cc, bcc)
        # 通知
        notify('Message has been sent to %s' % ', '.join(to))

    def convert(self, item):
        # extract filename & timestamp
        filename = re.match('^<?(.*?)>?$', item['Message-Id']).group(1)
        timestamp = email.utils.mktime_tz(email.utils.parsedate_tz(item['Date']))
        # date - self.convert by strftime accrding to format string defined as 30902
        parsed = email.utils.parsedate_tz(item['Date'])
        t = email.utils.mktime_tz(parsed)
        item['Date'] = time.strftime(self.addon.getLocalizedString(30902), time.localtime(t))
        # body - extract plain text from html using beautifulsoup
        decoded_body = item['body']
        if re.compile(r'<html|<!doctype', re.IGNORECASE).match(decoded_body):
            decoded_body = re.compile(r'<style.*?</style>', re.DOTALL|re.MULTILINE|re.IGNORECASE).sub('',decoded_body)
            decoded_body = re.compile(r'<script.*?</script>', re.DOTALL|re.MULTILINE|re.IGNORECASE).sub('',decoded_body)
            soup = BeautifulSoup(decoded_body, 'html.parser')
            buf = []
            for text in soup.stripped_strings:
                buf.append(text)
            item['body'] = ' '.join(buf)
        # 文字コード変換
        for key in item.keys():
            if isinstance(item[key], unicode): item[key] = item[key].encode('utf-8')
        return filename, timestamp

    def receive(self, criterion):
        # メール取得
        mails = self.service.receive(criterion, 'TEXT')
        # 新規メールのリスト
        newmails = []
        senders = {}
        # メール毎の処理
        for mail in mails:
            # データ変換
            filename, timestamp = self.convert(mail)
            # ファイルパス
            filepath = os.path.join(self.cache_path, filename)
            if not os.path.isfile(filepath):
                content = ''
                # header
                for key in ['From','To','CC','Date','Subject','Message-Id']:
                    if mail[key]: content += '%s: %s\r\n' % (key, mail[key].replace('\r\n',''))
                # separator
                content += '\r\n'
                # body
                content += mail['body']
                # ファイル書き込み
                f = open(filepath,'w')
                f.write(content)
                f.close()
                # タイムスタンプを変更
                os.utime(filepath, (timestamp, timestamp))
                # 新規メールのパスを格納
                newmails.append(filepath)
                senders[filepath] = mail['From'].replace('\r\n','')
        log('%d/%d mails retrieved using criterion "%s"' % (len(newmails),len(mails),criterion))
        return newmails, senders

    def list(self, newmails):
        #ファイルリスト
        files = sorted(os.listdir(self.cache_path), key=lambda mail: os.stat(os.path.join(self.cache_path,mail)).st_mtime, reverse=True)
        #ファイルの情報を表示
        count = 0
        for filename in files:
            filepath = os.path.join(self.cache_path,filename)
            if os.path.isfile(filepath) and not filename.startswith('.'):
                if self.listsize == 0 or count < self.listsize:
                    # ファイル読み込み
                    f = open(filepath,'r')
                    lines = f.readlines();
                    f.close()
                    # パース
                    params = {'Subject':'', 'From':'', 'body':False, 'messages':[]}
                    for line in lines:
                        line = line.rstrip()
                        if line == '':
                            params['body'] = True
                        elif params['body'] == True:
                            if line.startswith('?') or line.startswith('？'):
                                message = line[1:].strip()
                                if message: params['messages'].append(message)
                        elif params['body'] == False:
                            pair = line.split(': ',1)
                            if len(pair) == 2:
                                params[pair[0]] = pair[1]
                    # GUI設定
                    if params['Subject']:
                        title = params['Subject']
                    else:
                        title = self.addon.getLocalizedString(30901).encode('utf-8')
                    if filepath in newmails:
                        title = '[COLOR yellow]%s[/COLOR]' % title
                    # 日付文字列
                    d = datetime.datetime.fromtimestamp(os.stat(filepath).st_mtime)
                    dayfmt = self.addon.getLocalizedString(30904).encode('utf-8')
                    daystr = self.addon.getLocalizedString(30905).encode('utf-8')
                    fd = formatted_datetime(d, dayfmt, daystr)
                    # リストアイテム
                    label = '%s  %s' % (fd, title)
                    listitem = xbmcgui.ListItem(label, iconImage="DefaultFile.png", thumbnailImage="DefaultFile.png")
                    query = '%s?action=open&filename=%s' % (sys.argv[0],urllib.quote_plus(filename))
                    # コンテクストメニュー
                    menu = []
                    # 新着確認
                    menu.append((self.addon.getLocalizedString(30801),'Container.Update(%s,replace)' % (sys.argv[0])))
                    # アドオン設定
                    menu.append((self.addon.getLocalizedString(30802),'Addon.OpenSettings(%s)' % (self.addon.getAddonInfo('id'))))
                    # 追加
                    listitem.addContextMenuItems(menu, replaceItems=True)
                    xbmcplugin.addDirectoryItem(int(sys.argv[1]), query, listitem, False)
                    # カウントアップ
                    count += 1
                else:
                    os.remove(filepath)
        # 表示対象がない場合
        if count == 0:
            title = '[COLOR gray]%s[/COLOR]' % self.addon.getLocalizedString(30903)
            listitem = xbmcgui.ListItem(title, iconImage="DefaultFolder.png", thumbnailImage="DefaultFolder.png")
            # コンテクストメニュー
            menu = []
            # 新着確認
            menu.append((self.addon.getLocalizedString(30801),'Container.Update(%s,replace)' % (sys.argv[0])))
            # アドオン設定
            menu.append((self.addon.getLocalizedString(30802),'Addon.OpenSettings(%s)' % (self.addon.getAddonInfo('id'))))
            # 追加
            listitem.addContextMenuItems(menu, replaceItems=True)
            xbmcplugin.addDirectoryItem(int(sys.argv[1]), '', listitem, False)
        xbmcplugin.endOfDirectory(int(sys.argv[1]))

    def open(self, filename):
        # ファイル読み込み
        filepath = os.path.join(self.cache_path, filename)
        f = open(filepath,'r')
        lines = f.readlines();
        f.close()
        # パース
        header = []
        body = []
        params = {'Subject':self.addon.getLocalizedString(30901), 'body':False}
        for line in lines:
            line = line.rstrip()
            if params['body'] == True:
                body.append(line)
            elif line == '':
                params['body'] = True
            else:
                pair = line.split(': ',1)
                if len(pair) == 2:
                    params[pair[0]] = pair[1]
                    if pair[0] != 'Subject' and pair[0] != 'Message-Id':
                        header.append('[COLOR green]%s:[/COLOR] %s' % (pair[0],pair[1]))
        # テキストビューア
        viewer_id = 10147
        # ウィンドウを開く
        xbmc.executebuiltin('ActivateWindow(%s)' % viewer_id)
        # ウィンドウの用意ができるまで1秒待つ
        xbmc.sleep(1000)
        # ウィンドウへ書き込む
        viewer = xbmcgui.Window(viewer_id)
        viewer.getControl(1).setLabel('[COLOR orange]%s[/COLOR]' % params['Subject'])
        viewer.getControl(5).setText('%s\n\n%s' % ('\n'.join(header),'\n'.join(body)))


if __name__  == '__main__':
    main = Main()
    if main.service:
        main.main()
    else:
        xbmcaddon.Addon().openSettings()
