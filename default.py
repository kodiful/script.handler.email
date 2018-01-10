# -*- coding: utf-8 -*-

from __future__ import unicode_literals

import xbmcgui
import xbmcplugin
import xbmcaddon

import base64
import os
import codecs
import urllib, urllib2
import re
import threading

import email.utils
import time, datetime

from bs4 import BeautifulSoup
from resources.lib.gmail import Gmail

# アドオン情報
addon_id = 'script.message.tv'
addon = xbmcaddon.Addon(addon_id)

# ファイル/ディレクトリパス
PROFILE_PATH = xbmc.translatePath(addon.getAddonInfo('profile').decode('utf-8'))
PLUGIN_PATH = xbmc.translatePath(addon.getAddonInfo('path').decode('utf-8'))
RESOURCES_PATH = os.path.join(PLUGIN_PATH, 'resources')
DATA_PATH = os.path.join(RESOURCES_PATH, 'data')
LIB_PATH = os.path.join(RESOURCES_PATH, 'lib')
CACHE_PATH = os.path.join(PROFILE_PATH, 'cache')
if not os.path.isdir(CACHE_PATH): os.makedirs(CACHE_PATH)

TONE_FILE = os.path.join(DATA_PATH, 'Audio003201.mp3')
TEMPLATE_FILE = os.path.join(DATA_PATH, 'template.txt')

# テキストビューア
viewer_id = 10147


def addon_debug(message, level=xbmc.LOGNOTICE):
    debug = addon.getSetting('debug')
    if debug == 'true': addon_error(message, level)

def addon_error(message, level=xbmc.LOGERROR):
    try: xbmc.log(message, level)
    except: xbmc.log(message.encode('utf-8','ignore'), level)

def sendMail(subject, message, to, cc=None, replyto=None):
    # アカウント情報
    service = addon.getSetting('service')
    address = addon.getSetting('address')
    password = addon.getSetting('password')
    # クラス
    gmail = Gmail(service, address, password)
    # メール送信
    gmail.send(subject, message, [to], cc, replyto_address=replyto)
    # 通知
    xbmc.executebuiltin('XBMC.Notification("Gmail","Message has been sent to %s",3000,"DefaultIconInfo.png")' % to)
    addon_debug('gmail:sendMail() - send message to %s' % to)

def convert(item):
    # extract filename & timestamp
    filename = re.match('^<?(.*?)>?$', item['Message-Id']).group(1)
    timestamp = email.utils.mktime_tz(email.utils.parsedate_tz(item['Date']))
    # date - convert by strftime accrding to format string defined as 30902
    parsed = email.utils.parsedate_tz(item['Date'])
    t = email.utils.mktime_tz(parsed)
    item['Date'] = time.strftime(addon.getLocalizedString(30902), time.localtime(t))
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


def recvMails(criterion):
    # アカウント情報
    service = addon.getSetting('service')
    address = addon.getSetting('address')
    password = addon.getSetting('password')
    # メール取得
    mails = Gmail(service, address, password).recv(criterion, 'TEXT')
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
    return mails


def recvMesgs(data):
    # メッセージID、ファイル名
    messageid = '%s@kodiful.com' % (data['time'] or str(time.time()))
    filepath = os.path.join(CACHE_PATH,messageid)
    # ファイルの重複をチェック
    if os.path.isfile(filepath):
        return False
    # アカウント情報
    service = addon.getSetting('service')
    address = addon.getSetting('address')
    password = addon.getSetting('password')
    gmail = Gmail(service, address, password)
    user = gmail.get_from_address()
    # 現在時刻
    date = email.utils.formatdate(localtime=True)
    timestamp = email.utils.mktime_tz(email.utils.parsedate_tz(date))
    # date - convert by strftime accrding to format string defined as 30902
    parsed = email.utils.parsedate_tz(date)
    t = email.utils.mktime_tz(parsed)
    date = time.strftime(addon.getLocalizedString(30902), time.localtime(t))
    # パラメータをデコード
    name = urllib.unquote_plus(data['name'].encode('utf-8'))
    addr = urllib.unquote_plus(data['from'].encode('utf-8'))
    if name and addr:
        sender = '%s <%s>'.encode('utf-8') % (name,addr)
    elif name:
        sender = '%s'.encode('utf-8') % (name)
    elif addr:
        sender = '%s'.encode('utf-8') % (addr)
    else:
        sender = ''
    subject = urllib.unquote_plus(data['subject'].encode('utf-8'))
    message = urllib.unquote_plus(data['message'].encode('utf-8'))
    # ファイル書き込み
    f = codecs.open(TEMPLATE_FILE,'r','utf-8')
    template = f.read().encode('utf-8')
    f.close()
    source = template.format(
        sender=sender,
        recipient=user,
        date=date,
        subject=subject,
        messageid=messageid,
        message=message
        ).decode('utf-8')
    f = codecs.open(filepath,'w','utf-8')
    f.write(source)
    f.close()
    # タイムスタンプを変更
    os.utime(filepath, (timestamp, timestamp))
    return True

def sendMesgs(data, cc_addresses):
    # アカウント情報
    service = addon.getSetting('service')
    address = addon.getSetting('address')
    password = addon.getSetting('password')
    # パラメータ
    name = urllib.unquote_plus(data['name'].encode('utf-8'))
    from_address = urllib.unquote_plus(data['from'].encode('utf-8'))
    subject = urllib.unquote_plus(data['subject'].encode('utf-8'))
    message = urllib.unquote_plus(data['message'].encode('utf-8'))
    # 同報送信
    gmail = Gmail(service, address, password)
    gmail.send(('%s - fowarded by kodiful.com' % (subject.decode('utf-8'))).encode('utf-8'), message, cc_addresses, replyto_address=from_address)


def listMails(handle):
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
                    title = addon.getLocalizedString(30901)
                if filename.endswith('@kodiful.com'):
                    title = '[COLOR orange]%s[/COLOR]' % (title)
                if addon.getSetting('instantreply') == 'true' and len(params['messages']) > 0:
                    title = '%s [COLOR orange]%s[/COLOR]' % (title,addon.getLocalizedString(30909))
                plot = '[COLOR green]Date:[/COLOR] %s\n[COLOR green]From:[/COLOR] %s' % (params['Date'],params['From'])
                query = '%s?action=open&filename=%s' % (sys.argv[0],urllib.quote_plus(filename))
                listitem = xbmcgui.ListItem(title, iconImage="DefaultFolder.png", thumbnailImage="DefaultFolder.png")
                listitem.setInfo(type='video', infoLabels={'title':title, 'plot':plot})
                # コンテクストメニュー
                menu = []
                # 返信メニュー
                if addon.getSetting('instantreply') == 'true':
                    subject = urllib.quote_plus(('Re: %s' % params['Subject']).encode('utf-8','replace'))
                    to = urllib.quote_plus(params['From'].encode('utf-8','replace'))
                    if len(params['messages']) > 0:
                        messages = params['messages']
                    else:
                        messages = addon.getLocalizedString(30906).split('|')
                    for message in messages:
                        action = 'XBMC.RunPlugin(plugin://%s/?action=reply&subject=%s&message=%s&to=%s)' % (addon_id,subject,urllib.quote_plus(message.encode('utf-8','replace')),to)
                        menu.append(('[COLOR orange]%s[/COLOR]'%message,action))
                # 新着確認
                menu.append((addon.getLocalizedString(30201),'XBMC.RunPlugin(plugin://%s/?action=)' % (addon_id)))
                # アドオン設定
                menu.append((addon.getLocalizedString(30202),'Addon.OpenSettings(%s)' % (addon_id)))
                # 追加
                listitem.addContextMenuItems(menu, replaceItems=True)
                xbmcplugin.addDirectoryItem(handle, query, listitem, False)
                # カウントアップ
                count += 1
            else:
                os.remove(filepath)
    # 表示対象がない場合
    if count == 0:
        title = '[COLOR gray]%s[/COLOR]' % addon.getLocalizedString(30903)
        listitem = xbmcgui.ListItem(title, iconImage="DefaultFolder.png", thumbnailImage="DefaultFolder.png")
        listitem.setInfo(type='video', infoLabels={'title':title})
        xbmcplugin.addDirectoryItem(handle, '', listitem, False)
    xbmcplugin.endOfDirectory(handle)
    return len(files) > 0 and files[0]


def updateMails(popup=False):
    # メールのリストを表示
    filename = listMails(int(sys.argv[1]))
    # フォーカス
    if filename:
        xbmc.sleep(1000)
        retry = 50
        while (retry > 0):
            try:
                retry -= 1
                win = xbmcgui.Window(xbmcgui.getCurrentWindowId())
                win.getControl(win.getFocusId()).selectItem(0)
                break
            except:
                addon_debug('gmail:updateMails() - retry remains %d' % retry)
                xbmc.sleep(100)
    # ポップアップ
    if filename and popup:
        # Kodiをアクティベート
        if addon.getSetting('cec') == 'true':
            xbmc.executebuiltin('XBMC.CECActivateSource')
            xbmc.sleep(3000)
        # メールの内容を表示
        openMail(filename)
        # トーンを奏鳴
        if addon.getSetting('tone') == 'true':
            xbmc.executebuiltin('PlayMedia(%s)' % TONE_FILE)


def openMail(filename):
    # ファイル読み込み
    filepath = os.path.join(CACHE_PATH, filename)
    f = codecs.open(filepath,'r','utf-8')
    lines = f.readlines();
    f.close()
    # パース
    header = []
    body = []
    params = {'Subject':addon.getLocalizedString(30901), 'body':False}
    for line in lines:
        line = line.rstrip()
        if params['body'] == True:
            if addon.getSetting('instantreply') == 'true' and (line.startswith('?') or line.startswith('？')):
                body.append('[COLOR orange]%s[/COLOR]'%line)
            else:
                body.append(line)
        elif line == '':
            params['body'] = True
        else:
            pair = line.split(': ',1)
            if len(pair) == 2:
                params[pair[0]] = pair[1]
                if pair[0] != 'Subject' and pair[0] != 'Message-Id':
                    header.append('[COLOR green]%s:[/COLOR] %s' % (pair[0],pair[1]))
    # 表示
    xbmc.executebuiltin('ActivateWindow(%d)' % viewer_id)
    xbmc.sleep(1000)
    try:
        viewer = xbmcgui.Window(viewer_id)
        viewer.getControl(1).setLabel('[COLOR orange]%s[/COLOR]' % params['Subject'])
        viewer.getControl(5).setText('%s\n\n%s' % ('\n'.join(header),'\n'.join(body)))
    except:
        addon_debug('gmail:openMail() - getControl failed')


def main():
    # ログ
    for arg in sys.argv:
        addon_debug('gmail:main() - argv: %s' % arg)
    # パラメータ抽出
    params = {'action':'','filename':'','subject':'','time':'','message':'','name':'','from':'', 'cc':''}
    if len(sys.argv[2]) > 0:
        pairs = re.compile(r'[?&]').split(sys.argv[2])
        for i in range(len(pairs)):
            splitted = pairs[i].split('=',1)
            if len(splitted) == 2:
                params[splitted[0]] = splitted[1]
    # ログ
    for p in params:
        addon_debug('gmail:main() - params[%s]: %s' % (p,params[p]))
    # メイン処理
    if addon.getSetting('address') == '' or addon.getSetting('password') == '':
        # アカウント情報設定
        addon.openSettings()
    elif params['action'] == '':
        # 新着メール取得
        mails = recvMails('UNSEEN')
        addon_debug('gmail:main() - %d mails retrieved' % len(mails))
        # 表示更新
        updateMails(len(mails)>0)
    elif params['action'] == 'refresh':
        # キャッシュクリア
        files = os.listdir(CACHE_PATH)
        for filename in files:
            try:
                if not filename.endswith('@kodiful.com'):
                    os.remove(os.path.join(CACHE_PATH, filename))
            except: pass
        # メール取得
        #mails = recvMails('SINCE 1-Jan-2015')
        d = datetime.datetime.now() - datetime.timedelta(days=30)
        criterion = d.strftime('SINCE %d-%b-%Y').decode('utf-8')
        mails = recvMails(criterion)
        addon_debug('gmail:main() - %d mails retrieved' % len(mails))
        # 表示更新
        updateMails(False)
    elif params['action'] == 'open':
        # メールの内容を表示
        filename = urllib.unquote_plus(params['filename'])
        if filename: openMail(filename)
    elif params['action'] == 'reply':
        # メールに定型文で返信
        subject = urllib.unquote_plus(str(params['subject']))
        message = urllib.unquote_plus(str(params['message']))
        to = urllib.unquote_plus(str(params['to']))
        sendMail(subject, message, to)
    elif params['action'] == 'login':
        pass
    elif params['action'] == 'message':
        # RPCメッセージ取得
        if recvMesgs(params):
            # 同報
            if addon.getSetting('cc') == 'true':
                cc_addresses = addon.getSetting('cc_addresses')
                if cc_addresses: sendMesgs(params, re.split('\s*,\s*',cc_addresses))
            # 表示更新
            updateMails(True)

if __name__  == '__main__': main()
