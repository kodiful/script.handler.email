# -*- coding: utf-8 -*-

import xbmc

from default import Main
from resources.lib.common import Common


class Monitor(xbmc.Monitor):

    def __init__(self, *args, **kwargs):
        self.timer = 1
        super().__init__()

    def check_timer(self):
        interval = Common.GET('interval')
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
        Common.log('settings changed')

    def onScreensaverActivated(self):
        Common.log('screensaver activated')

    def onScreensaverDeactivated(self):
        Common.log('screensaver deactivated')


if __name__ == "__main__":
    monitor = Monitor()
    while not monitor.abortRequested():
        if monitor.waitForAbort(60):
            break
        if monitor.check_timer():
            main = Main()
            if main.service:
                main.check(refresh=False)
            else:
                Common.ADDON.openSettings()
