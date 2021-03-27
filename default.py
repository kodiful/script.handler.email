# -*- coding: utf-8 -*-

import xbmc
import xbmcgui
import xbmcplugin
import xbmcaddon
import xbmcvfs

import os
import re
import sys
import email.utils
import time
import datetime
import locale

from urllib.parse import parse_qs
from urllib.parse import urlencode
from urllib.parse import quote
from bs4 import BeautifulSoup

from resources.lib.common import Common
from resources.lib.mail import Mail
from resources.lib.gmail import Gmail
from resources.lib.icloud import iCloud


class Main:

    def __init__(self, service=None):
        # メールサービスを初期化
        self.service = None
        service = service or Common.GET('service')
        if service == 'Custom':
            user = Common.GET('user')
            password = Common.GET('password')
            smtp_host = Common.GET('smtp_host')
            smtp_port = Common.GET('smtp_port')
            smtp_auth = Common.GET('smtp_auth')
            smtp_ssl = Common.GET('smtp_ssl')
            smtp_tls = Common.GET('smtp_tls')
            smtp_from = Common.GET('smtp_from')
            imap_host = Common.GET('imap_host')
            imap_port = Common.GET('imap_port')
            imap_ssl = Common.GET('imap_ssl')
            imap_tls = Common.GET('imap_tls')
            if user and password and smtp_host and smtp_port and smtp_from and smtp_host and smtp_port:
                config = {
                    'smtp_host': smtp_host,
                    'smtp_port': int(smtp_port),
                    'smtp_auth': smtp_auth == 'true',
                    'smtp_ssl': smtp_ssl == 'true',
                    'smtp_tls': smtp_tls == 'true',
                    'smtp_from': smtp_from,
                    'imap_host': imap_host,
                    'imap_port': int(imap_port),
                    'imap_ssl': imap_ssl == 'true',
                    'imap_tls': imap_tls == 'true'
                }
                self.service = Mail(user, password, config)
        elif service == 'Gmail':
            user = Common.GET('user1')
            password = Common.GET('password1')
            if user and password:
                self.service = Gmail(user, password)
        elif service == 'iCloud':
            user = Common.GET('user2')
            password = Common.GET('password2')
            if user and password:
                self.service = iCloud(user, password)
        # メールサービスが正常に初期化されたら他の初期化を実行
        if self.service:
            # キャッシュディレクトリのパス
            profile_path = xbmcvfs.translatePath(Common.INFO('profile'))
            self.cache_path = os.path.join(profile_path, 'cache', service)
            if not os.path.isdir(self.cache_path):
                os.makedirs(self.cache_path)
            # 表示するメール数
            self.listsize = Common.GET('listsize')
            if self.listsize == 'Unlimited':
                self.listsize = 0
            else:
                self.listsize = int(self.listsize)

    def main(self):
        # パラメータ抽出
        params = {'action': '', 'filename': '', 'subject': '', 'message': '', 'to': [], 'cc': []}
        args = parse_qs(sys.argv[2][1:])
        for key in params.keys():
            params[key] = args.get(key, params[key])
        for key in ['action', 'filename', 'subject', 'message']:
            params[key] = params[key] and params[key][0]
        # メイン処理
        if params['action'] == '':
            start = Common.GET('start')
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
            if params['filename']:
                self.open(params['filename'])
        elif params['action'] == 'sendmessage':
            # メールを送信
            subject = Common.GET('subject')
            message = Common.GET('message')
            to = Common.GET('to')
            cc = Common.GET('cc')
            # 送信データ
            values = {'action': 'send', 'subject': subject, 'message': message, 'to': to, 'cc': cc}
            postdata = urlencode(values)
            xbmc.executebuiltin('RunPlugin(%s?%s)' % (sys.argv[0], postdata))
        elif params['action'] == 'prepmessage':
            Common.SET('subject', params['subject'])
            Common.SET('message', params['message'])
            Common.SET('to', ','.join(params['to']))
            Common.SET('cc', ','.join(params['cc']))
            xbmc.executebuiltin('Addon.OpenSettings(%s)' % Common.ADDON_ID)
            xbmc.executebuiltin('SetFocus(101)')  # 2nd category
            xbmc.executebuiltin('SetFocus(203)')  # 4th control
        elif params['action'] == 'send':
            # メールを送信
            if Common.GET('bcc') == 'true':
                bcc = [self.service.smtp_from]
            else:
                bcc = []
            self.send(subject=params['subject'], message=params['message'], to=params['to'], cc=params['cc'], bcc=bcc)

    def check(self, refresh=True):
        # 管理用ファイル
        criterion_file = os.path.join(self.cache_path, '.criterion')
        newmails_file = os.path.join(self.cache_path, '.newmails')
        # ロケールを変更
        locale.setlocale(locale.LC_TIME, 'en_US.UTF-8')
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
            if Common.GET('cec') == 'true':
                xbmc.executebuiltin('XBMC.CECActivateSource')
            # 新着があることを通知
            Common.notify('New mail from %s' % senders[newmails[0]])
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
        Common.notify('Message has been sent to %s' % ', '.join(to))

    def convert(self, item):
        # extract filename & timestamp
        filename = re.match('^<?(.*?)>?$', item['Message-Id']).group(1)
        timestamp = email.utils.mktime_tz(email.utils.parsedate_tz(item['Date']))
        # date - self.convert by strftime accrding to format string defined as 30902
        parsed = email.utils.parsedate_tz(item['Date'])
        t = email.utils.mktime_tz(parsed)
        item['Date'] = time.strftime(Common.STR(30902), time.localtime(t))
        # body - extract plain text from html using beautifulsoup
        decoded_body = item['body']
        if re.compile(r'<html|<!doctype', re.IGNORECASE).match(decoded_body):
            decoded_body = re.compile(r'<style.*?</style>', re.DOTALL | re.MULTILINE | re.IGNORECASE).sub('', decoded_body)
            decoded_body = re.compile(r'<script.*?</script>', re.DOTALL | re.MULTILINE | re.IGNORECASE).sub('', decoded_body)
            soup = BeautifulSoup(decoded_body, 'html.parser')
            buf = []
            for text in soup.stripped_strings:
                buf.append(text)
            item['body'] = ' '.join(buf)
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
                for key in ['From', 'To', 'CC', 'Date', 'Subject', 'Message-Id']:
                    if mail[key]:
                        content += '%s: %s\r\n' % (key, mail[key].replace('\r\n', ''))
                # separator
                content += '\r\n'
                # body
                content += mail['body']
                # ファイル書き込み
                f = open(filepath, 'w')
                f.write(content)
                f.close()
                # タイムスタンプを変更
                os.utime(filepath, (timestamp, timestamp))
                # 新規メールのパスを格納
                newmails.append(filepath)
                senders[filepath] = mail['From'].replace('\r\n', '')
        Common.log('%d/%d mails retrieved using criterion "%s"' % (len(newmails), len(mails), criterion))
        return newmails, senders

    def list(self, newmails):
        # ファイルリスト
        files = sorted(os.listdir(self.cache_path), key=lambda mail: os.stat(os.path.join(self.cache_path, mail)).st_mtime, reverse=True)
        # ファイルの情報を表示
        count = 0
        for filename in files:
            filepath = os.path.join(self.cache_path, filename)
            if os.path.isfile(filepath) and not filename.startswith('.'):
                if self.listsize == 0 or count < self.listsize:
                    # ファイル読み込み
                    f = open(filepath, 'r')
                    lines = f.readlines()
                    f.close()
                    # パース
                    params = {'Subject': '', 'From': '', 'body': False, 'messages': []}
                    for line in lines:
                        line = line.rstrip()
                        if line == '':
                            params['body'] = True
                        elif params['body'] is True:
                            if line.startswith('?') or line.startswith('？'):
                                message = line[1:].strip()
                                if message:
                                    params['messages'].append(message)
                        elif params['body'] is False:
                            pair = line.split(': ', 1)
                            if len(pair) == 2:
                                params[pair[0]] = pair[1]
                    # GUI設定
                    if params['Subject']:
                        title = params['Subject']
                    else:
                        title = Common.STR(30901)
                    if filepath in newmails:
                        title = '[COLOR yellow]%s[/COLOR]' % title
                    # 日付文字列
                    d = datetime.datetime.fromtimestamp(os.stat(filepath).st_mtime)
                    # リストアイテム
                    label = '%s  %s' % (Common.datetime(d), title)
                    listitem = xbmcgui.ListItem(label)
                    listitem.setArt({'icon': 'DefaultFile.png', 'thumb': 'DefaultFile.png'})
                    query = '%s?action=open&filename=%s' % (sys.argv[0], quote(filename))
                    # コンテクストメニュー
                    menu = []
                    # 返信
                    to = params.get('From', '')
                    cc = params.get('CC', '')
                    subject = params.get('Subject', '')
                    values = {'action': 'prepmessage', 'subject': 'Re: %s' % subject, 'message': '', 'to': to, 'cc': cc}
                    postdata = urlencode(values)
                    menu.append((Common.STR(30800), 'RunPlugin(%s?%s)' % (sys.argv[0], postdata)))
                    # 新着確認
                    menu.append((Common.STR(30801), 'Container.Update(%s,replace)' % (sys.argv[0])))
                    # アドオン設定
                    menu.append((Common.STR(30802), 'Addon.OpenSettings(%s)' % (Common.ADDON_ID)))
                    # 追加
                    listitem.addContextMenuItems(menu, replaceItems=True)
                    xbmcplugin.addDirectoryItem(int(sys.argv[1]), query, listitem, False)
                    # カウントアップ
                    count += 1
                else:
                    os.remove(filepath)
        # 表示対象がない場合
        if count == 0:
            title = '[COLOR gray]%s[/COLOR]' % Common.STR(30903)
            listitem = xbmcgui.ListItem(title)
            listitem.setArt({'icon': 'DefaultFolder.png', 'thumb': 'DefaultFolder.png'})
            # コンテクストメニュー
            menu = []
            # 新着確認
            menu.append((Common.STR(30801), 'Container.Update(%s,replace)' % (sys.argv[0])))
            # アドオン設定
            menu.append((Common.STR(30802), 'Addon.OpenSettings(%s)' % (Common.ADDON_ID)))
            # 追加
            listitem.addContextMenuItems(menu, replaceItems=True)
            xbmcplugin.addDirectoryItem(int(sys.argv[1]), '', listitem, False)
        xbmcplugin.endOfDirectory(int(sys.argv[1]))

    def open(self, filename):
        # ファイル読み込み
        filepath = os.path.join(self.cache_path, filename)
        f = open(filepath, 'r')
        lines = f.readlines()
        f.close()
        # パース
        header = []
        body = []
        params = {'Subject': Common.STR(30901), 'body': False}
        for line in lines:
            line = line.rstrip()
            if params['body'] is True:
                body.append(line)
            elif line == '':
                params['body'] = True
            else:
                pair = line.split(': ', 1)
                if len(pair) == 2:
                    params[pair[0]] = pair[1]
                    if pair[0] != 'Subject' and pair[0] != 'Message-Id':
                        header.append('[COLOR green]%s:[/COLOR] %s' % (pair[0], pair[1]))
        # テキストビューア
        viewer_id = 10147
        # ウィンドウを開く
        xbmc.executebuiltin('ActivateWindow(%s)' % viewer_id)
        # ウィンドウの用意ができるまで1秒待つ
        xbmc.sleep(1000)
        # ウィンドウへ書き込む
        viewer = xbmcgui.Window(viewer_id)
        viewer.getControl(1).setLabel('[COLOR orange]%s[/COLOR]' % params['Subject'])
        viewer.getControl(5).setText('%s\n\n%s' % ('\n'.join(header), '\n'.join(body)))


if __name__ == '__main__':
    main = Main()
    if main.service:
        main.main()
    else:
        xbmcaddon.Addon().openSettings()
