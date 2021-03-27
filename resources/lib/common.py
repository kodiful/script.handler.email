# -*- coding: utf-8 -*-

from resources.lib.commonmethods import Common as C


class Common(C):

    @staticmethod
    def datetime(d):
        weekday = d.weekday()
        weekdaystr = Common.STR(30905).split(',')[weekday]
        datetimestr = d.strftime('%Y-%m-%d(%%s) %H:%M:%S') % weekdaystr
        if weekday == 6 or Common.isholiday(d.strftime('%Y-%m-%d')):
            template = '[COLOR red]%s[/COLOR]'
        elif weekday == 5:
            template = '[COLOR blue]%s[/COLOR]'
        else:
            template = '%s'
        return template % datetimestr
