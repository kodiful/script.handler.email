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

# アドオン情報
ADDON = xbmcaddon.Addon()

# ファイル/ディレクトリパス
PROFILE_PATH = xbmc.translatePath(ADDON.getAddonInfo('profile').decode('utf-8'))
PLUGIN_PATH = xbmc.translatePath(ADDON.getAddonInfo('path').decode('utf-8'))
RESOURCES_PATH = os.path.join(PLUGIN_PATH, 'resources')
DATA_PATH = os.path.join(RESOURCES_PATH, 'data')
LIB_PATH = os.path.join(RESOURCES_PATH, 'lib')
CACHE_PATH = os.path.join(PROFILE_PATH, 'cache')
if not os.path.isdir(CACHE_PATH): os.makedirs(CACHE_PATH)
TEMPLATE_FILE = os.path.join(DATA_PATH, 'template.txt')

# テキストビューア
viewer_id = 10147

def sendMail(subject, message, to, cc=None, replyto=None):
    # アカウント情報
    service = ADDON.getSetting('service')
    address = ADDON.getSetting('address')
    password = ADDON.getSetting('password')
    # クラス
    gmail = Gmail(service, address, password)
    # メール送信
    gmail.send(subject, message, [to], cc, replyto_address=replyto)
    # 通知
    notify('Message has been sent to %s' % to)
    log('send message to %s' % to)

def convert(item):
    # extract filename & timestamp
    filename = re.match('^<?(.*?)>?$', item['Message-Id']).group(1)
    timestamp = email.utils.mktime_tz(email.utils.parsedate_tz(item['Date']))
    # date - convert by strftime accrding to format string defined as 30902
    parsed = email.utils.parsedate_tz(item['Date'])
    t = email.utils.mktime_tz(parsed)
    item['Date'] = time.strftime(ADDON.getLocalizedString(30902), time.localtime(t))
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

def receiveMails(criterion):
    # アカウント情報
    service = ADDON.getSetting('service')
    address = ADDON.getSetting('address')
    password = ADDON.getSetting('password')
    # メール取得
    mails = Gmail(service, address, password).receive(criterion, 'TEXT')
    # メール毎の処理
    for mail in mails:
        # データ変換
        filename, timestamp = convert(mail)
        # ファイル書き込み
        filepath = os.path.join(CACHE_PATH, filename)
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

def listMails():
    #ファイルリスト
    files = sorted(os.listdir(CACHE_PATH), key=lambda mail: os.stat(os.path.join(CACHE_PATH,mail)).st_mtime, reverse=True)
    #ファイルの情報を表示
    count = 0
    for filename in files:
        filepath = os.path.join(CACHE_PATH,filename)
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
                    title = ADDON.getLocalizedString(30901)
                plot = '[COLOR green]Date:[/COLOR] %s\n[COLOR green]From:[/COLOR] %s' % (params['Date'],params['From'])
                query = '%s?action=open&filename=%s' % (sys.argv[0],urllib.quote_plus(filename))
                listitem = xbmcgui.ListItem(title, iconImage="DefaultFolder.png", thumbnailImage="DefaultFolder.png")
                listitem.setInfo(type='video', infoLabels={'title':title, 'plot':plot})
                # コンテクストメニュー
                menu = []
                # 新着確認
                menu.append((ADDON.getLocalizedString(30201),'XBMC.RunPlugin(plugin://%s/?action=)' % (ADDON.getAddonInfo('id'))))
                # アドオン設定
                menu.append((ADDON.getLocalizedString(30202),'Addon.OpenSettings(%s)' % (ADDON.getAddonInfo('id'))))
                # 追加
                listitem.addContextMenuItems(menu, replaceItems=True)
                xbmcplugin.addDirectoryItem(int(sys.argv[1]), query, listitem, False)
                # カウントアップ
                count += 1
            else:
                os.remove(filepath)
    # 表示対象がない場合
    if count == 0:
        title = '[COLOR gray]%s[/COLOR]' % ADDON.getLocalizedString(30903)
        listitem = xbmcgui.ListItem(title, iconImage="DefaultFolder.png", thumbnailImage="DefaultFolder.png")
        listitem.setInfo(type='video', infoLabels={'title':title})
        xbmcplugin.addDirectoryItem(int(sys.argv[1]), '', listitem, False)
    xbmcplugin.endOfDirectory(int(sys.argv[1]))

def openMail(filename):
    # ファイル読み込み
    filepath = os.path.join(CACHE_PATH, filename)
    f = codecs.open(filepath,'r','utf-8')
    lines = f.readlines();
    f.close()
    # パース
    header = []
    body = []
    params = {'Subject':ADDON.getLocalizedString(30901), 'body':False}
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
    # ウィンドウを開く
    xbmc.executebuiltin('ActivateWindow(%d)' % viewer_id)
    # ウィンドウの用意ができるまで1秒待つ
    xbmc.sleep(1000)
    # ウィンドウへ書き込む
    viewer = xbmcgui.Window(viewer_id)
    viewer.getControl(1).setLabel('[COLOR orange]%s[/COLOR]' % params['Subject'])
    viewer.getControl(5).setText('%s\n\n%s' % ('\n'.join(header),'\n'.join(body)))

def main():
    # パラメータ抽出
    params = {'action':'','filename':'','subject':'','time':'','message':'','name':'','from':'', 'cc':''}
    args = urlparse.parse_qs(sys.argv[2][1:])
    for key in params.keys():
        val = args.get(key, None)
        if val: params[key] = val[0]
    # メイン処理
    if ADDON.getSetting('address') == '' or ADDON.getSetting('password') == '':
        # アカウント情報設定
        ADDON.openSettings()
    elif params['action'] == '':
        # 新着メール取得
        receiveMails('UNSEEN')
        listMails()
    elif params['action'] == 'refresh':
        # キャッシュクリア
        files = os.listdir(CACHE_PATH)
        for filename in files:
            os.remove(os.path.join(CACHE_PATH, filename))
        # メール取得
        #mails = receiveMails('SINCE 1-Jan-2015')
        d = datetime.datetime.now() - datetime.timedelta(days=30)
        criterion = d.strftime('SINCE %d-%b-%Y').decode('utf-8')
        receiveMails(criterion)
        listMails()
    elif params['action'] == 'open':
        # メールの内容を表示
        filename = urllib.unquote_plus(params['filename'])
        if filename: openMail(filename)

if __name__  == '__main__': main()
