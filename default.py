# -*- coding: utf-8 -*-

from __future__ import unicode_literals

import xbmcgui
import xbmcplugin
import xbmcaddon

import base64
import os
import codecs
import urllib, urllib2, urlparse
import re

import email.utils
import time, datetime

from bs4 import BeautifulSoup
from resources.lib.gmail import Gmail
from resources.lib.common import log, notify

class Main:

    def __init__(self):
        # アドオン
        self.addon = xbmcaddon.Addon()
        # ファイル/ディレクトリパス
        self.profile_path = xbmc.translatePath(self.addon.getAddonInfo('profile').decode('utf-8'))
        self.plugin_path = xbmc.translatePath(self.addon.getAddonInfo('path').decode('utf-8'))
        self.resources_path = os.path.join(self.plugin_path, 'resources')
        self.data_path = os.path.join(self.resources_path, 'data')
        self.lib_path = os.path.join(self.resources_path, 'lib')
        self.cache_path = os.path.join(self.profile_path, 'cache')
        self.template_file = os.path.join(self.data_path, 'template.txt')
        if not os.path.isdir(self.cache_path):
            os.makedirs(self.cache_path)

        # パラメータ抽出
        params = {'action':'','filename':'','subject':'','time':'','message':'','name':'','from':'', 'cc':''}
        args = urlparse.parse_qs(sys.argv[2][1:])
        for key in params.keys():
            val = args.get(key, None)
            if val: params[key] = val[0]
        # メイン処理
        if self.addon.getSetting('address') == '' or self.addon.getSetting('password') == '':
            # アカウント情報設定
            self.addon.openSettings()
        elif params['action'] == '':
            # 新着メール取得
            self.receive('UNSEEN')
            self.list()
        elif params['action'] == 'refresh':
            # キャッシュクリア
            files = os.listdir(self.cache_path)
            for filename in files:
                os.remove(os.path.join(self.cache_path, filename))
            # メール取得
            #mails = self.receive('SINCE 1-Jan-2015')
            d = datetime.datetime.now() - datetime.timedelta(days=30)
            criterion = d.strftime('SINCE %d-%b-%Y').decode('utf-8')
            self.receive(criterion)
            self.list()
        elif params['action'] == 'open':
            # メールの内容を表示
            filename = urllib.unquote_plus(params['filename'])
            if filename: self.open(filename)

    def send(self, subject, message, to, cc=None, replyto=None):
        # アカウント情報
        service = self.addon.getSetting('service')
        address = self.addon.getSetting('address')
        password = self.addon.getSetting('password')
        # クラス
        gmail = Gmail(service, address, password)
        # メール送信
        gmail.send(subject, message, [to], cc, replyto_address=replyto)
        # 通知
        notify('Message has been sent to %s' % to)
        log('send message to %s' % to)

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
            soup = BeautifulSoup(decoded_body)
            buf = []
            for text in soup.stripped_strings:
                buf.append(text)
            item['body'] = ' '.join(buf)
        return filename, timestamp

    def receive(self, criterion):
        # アカウント情報
        service = self.addon.getSetting('service')
        address = self.addon.getSetting('address')
        password = self.addon.getSetting('password')
        # メール取得
        mails = Gmail(service, address, password).receive(criterion, 'TEXT')
        # メール毎の処理
        for mail in mails:
            # データ変換
            filename, timestamp = self.convert(mail)
            # ファイル書き込み
            filepath = os.path.join(self.cache_path, filename)
            f = codecs.open(filepath,'w','utf-8')
            # header
            for key in ['From','To','CC','Date','Subject','Message-Id']:
                if mail[key]:
                    f.write('%s: ' % key)
                    f.write(mail[key].replace('\r\n',''))
                    f.write('\r\n')
            # separator
            f.write('\r\n')
            # body
            f.write(mail['body'])
            # ファイル書き込み完了
            f.close()
            # タイムスタンプを変更
            os.utime(filepath, (timestamp, timestamp))

    def list(self):
        #ファイルリスト
        files = sorted(os.listdir(self.cache_path), key=lambda mail: os.stat(os.path.join(self.cache_path,mail)).st_mtime, reverse=True)
        #ファイルの情報を表示
        count = 0
        for filename in files:
            filepath = os.path.join(self.cache_path,filename)
            if os.path.isfile(filepath) and not filename.startswith('.'):
                if count < 200:
                    # ファイル読み込み
                    f = codecs.open(filepath,'r','utf-8')
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
                        title = self.addon.getLocalizedString(30901)
                    plot = '[COLOR green]Date:[/COLOR] %s\n[COLOR green]From:[/COLOR] %s' % (params['Date'],params['From'])
                    query = '%s?action=open&filename=%s' % (sys.argv[0],urllib.quote_plus(filename))
                    listitem = xbmcgui.ListItem(title, iconImage="DefaultFolder.png", thumbnailImage="DefaultFolder.png")
                    listitem.setInfo(type='video', infoLabels={'title':title, 'plot':plot})
                    # コンテクストメニュー
                    menu = []
                    # 新着確認
                    menu.append((self.addon.getLocalizedString(30201),'XBMC.RunPlugin(plugin://%s/?action=)' % (self.addon.getAddonInfo('id'))))
                    # アドオン設定
                    menu.append((self.addon.getLocalizedString(30202),'Addon.OpenSettings(%s)' % (self.addon.getAddonInfo('id'))))
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
            listitem.setInfo(type='video', infoLabels={'title':title})
            xbmcplugin.addDirectoryItem(int(sys.argv[1]), '', listitem, False)
        xbmcplugin.endOfDirectory(int(sys.argv[1]))

    def open(self, filename):
        # ファイル読み込み
        filepath = os.path.join(self.cache_path, filename)
        f = codecs.open(filepath,'r','utf-8')
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


if __name__  == '__main__': Main()
