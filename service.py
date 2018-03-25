# -*- coding: utf-8 -*-

import xbmc, xbmcaddon

from default import Main
from resources.lib.common import log, notify

class Monitor(xbmc.Monitor):

    def __init__(self, *args, **kwargs):
        self.addon = xbmcaddon.Addon()
        self.timer = 1
        xbmc.Monitor.__init__(self)

    def check_timer(self):
        interval = self.addon.getSetting('interval')
        status = False
        if interval == 'None':
            self.timer = 0
        elif self.timer < int(interval):
            self.timer += 1
        else:
            status = True
            self.timer = 0
        return status

    def onSettingsChanged(self):
        log('settings changed')

    def onScreensaverActivated(self):
        log('screensaver activated')

    def onScreensaverDeactivated(self):
        log('screensaver deactivated')


if __name__ == "__main__":
    notify('Starting service', time=3000)
    monitor = Monitor()
    while not monitor.abortRequested():
        if monitor.waitForAbort(60):
            break
        if monitor.check_timer():
            main = Main()
            if main.service:
                main.check(refresh=False)
            else:
                xbmcaddon.Addon().openSettings()
