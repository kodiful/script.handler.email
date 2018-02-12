# -*- coding: utf-8 -*-

import inspect
import xbmc, xbmcaddon

def notify(message, **options):
    time = options.get('time', 10000)
    image = options.get('image', None)
    if image is None:
        if options.get('error', False):
            image = 'DefaultIconError.png'
        else:
            image = 'DefaultIconInfo.png'
    log(message, error=options.get('error', False))
    xbmc.executebuiltin('Notification("%s","%s",%d,"%s")' % (xbmcaddon.Addon().getAddonInfo('name'),message,time,image))

def log(*messages, **options):
    addon = xbmcaddon.Addon()
    if options.get('error', False):
        level = xbmc.LOGERROR
    elif addon.getSetting('debug') == 'true':
        level = xbmc.LOGNOTICE
    else:
        level = None
    if level:
        m = []
        for message in messages:
            if isinstance(message, str):
                m.append(message)
            elif isinstance(message, unicode):
                m.append(message.encode('utf-8'))
            else:
                m.append(str(message))
        frame = inspect.currentframe(1)
        xbmc.log(str('%s: %s %d: %s') % (addon.getAddonInfo('id'), frame.f_code.co_name, frame.f_lineno, str('').join(m)), level)

def formatted_datetime(d, dayfmt, daystr):
    if isinstance(dayfmt, str): dayfmt = dayfmt.decode('utf-8')
    if isinstance(daystr, str): daystr = daystr.decode('utf-8')
    weekday = d.weekday()
    weekdaystr = daystr.split(',')[weekday]
    date1 = d.strftime(dayfmt.encode('utf-8','ignore')).decode('utf-8') + '(' + weekdaystr + ')'
    time1 = d.strftime('%H:%M:%S').decode('utf-8')
    if weekday == 6 or isholiday(d.strftime('%Y-%m-%d')):
        template = u'[COLOR red]%s %s[/COLOR]'
    elif weekday == 5:
        template = u'[COLOR blue]%s %s[/COLOR]'
    else:
        template = u'%s %s'
    return template % (date1, time1)

def isholiday(day):
    holidays = {
        "2014-01-01":True,
        "2014-01-13":True,
        "2014-02-11":True,
        "2014-03-21":True,
        "2014-04-29":True,
        "2014-05-03":True,
        "2014-05-04":True,
        "2014-05-05":True,
        "2014-05-06":True,
        "2014-07-21":True,
        "2014-09-15":True,
        "2014-09-23":True,
        "2014-10-13":True,
        "2014-11-03":True,
        "2014-11-23":True,
        "2014-11-24":True,
        "2014-12-23":True,
        "2015-01-01":True,
        "2015-01-12":True,
        "2015-02-11":True,
        "2015-03-21":True,
        "2015-04-29":True,
        "2015-05-03":True,
        "2015-05-04":True,
        "2015-05-05":True,
        "2015-05-06":True,
        "2015-07-20":True,
        "2015-09-21":True,
        "2015-09-22":True,
        "2015-09-23":True,
        "2015-10-12":True,
        "2015-11-03":True,
        "2015-11-23":True,
        "2015-12-23":True,
        "2016-01-01":True,
        "2016-01-11":True,
        "2016-02-11":True,
        "2016-03-20":True,
        "2016-03-21":True,
        "2016-04-29":True,
        "2016-05-03":True,
        "2016-05-04":True,
        "2016-05-05":True,
        "2016-07-18":True,
        "2016-08-11":True,
        "2016-09-19":True,
        "2016-09-22":True,
        "2016-10-10":True,
        "2016-11-03":True,
        "2016-11-23":True,
        "2016-12-23":True,
        "2017-01-01":True,
        "2017-01-02":True,
        "2017-01-09":True,
        "2017-02-11":True,
        "2017-03-20":True,
        "2017-04-29":True,
        "2017-05-03":True,
        "2017-05-04":True,
        "2017-05-05":True,
        "2017-07-17":True,
        "2017-08-11":True,
        "2017-09-18":True,
        "2017-09-23":True,
        "2017-10-09":True,
        "2017-11-03":True,
        "2017-11-23":True,
        "2017-12-23":True,
        "2018-01-01":True,
        "2018-01-08":True,
        "2018-02-11":True,
        "2018-02-12":True,
        "2018-03-21":True,
        "2018-04-29":True,
        "2018-05-03":True,
        "2018-05-04":True,
        "2018-05-05":True,
        "2018-07-16":True,
        "2018-08-11":True,
        "2018-09-17":True,
        "2018-09-23":True,
        "2018-09-24":True,
        "2018-10-08":True,
        "2018-11-03":True,
        "2018-11-23":True,
        "2018-12-23":True,
        "2018-12-24":True,
        "2019-01-01":True,
        "2019-01-14":True,
        "2019-02-11":True,
        "2019-03-21":True,
        "2019-04-29":True,
        "2019-05-03":True,
        "2019-05-04":True,
        "2019-05-05":True,
        "2019-05-06":True,
        "2019-07-15":True,
        "2019-08-11":True,
        "2019-08-12":True,
        "2019-09-16":True,
        "2019-09-23":True,
        "2019-10-14":True,
        "2019-11-03":True,
        "2019-11-04":True,
        "2019-11-23":True,
        "2019-12-23":True,
        "2020-01-01":True,
        "2020-01-13":True,
        "2020-02-11":True,
        "2020-03-20":True,
        "2020-04-29":True,
        "2020-05-03":True,
        "2020-05-04":True,
        "2020-05-05":True,
        "2020-05-06":True,
        "2020-07-20":True,
        "2020-08-11":True,
        "2020-09-21":True,
        "2020-09-22":True,
        "2020-10-12":True,
        "2020-11-03":True,
        "2020-11-23":True,
        "2020-12-23":True}
    return holidays.get(day,False)
