import time
import random
from configs.logger import logger
from activities.apps.browsers.base import Browser

class GoogleChrome(Browser):
    def __init__(self):
        super().__init__()
        self.window_info = "[TITLE:New Tab - Google Chrome; CLASS:Chrome_WidgetWin_1]"

    def create_window(self):
        logger.info("Creating new Google Chrome window")
        if not self.dll.AU3_Run("C:\\Program Files\\Google\Chrome\\Application\\chrome.exe", "", 1):
            return "could not run winword.exe"
                
        time.sleep(2)
        logger.info("Checking existing Google Chrome window")
        if self.dll.AU3_WinExists(self.window_info, ""):
            logger.info("Getting Google Chrome window handle")
            handle = self.dll.AU3_WinGetHandle(self.window_info, "")
            if not handle:
                return "could not get window handle"
            else:
                self.window_info = "[HANDLE:%s]" % f"{handle:08X}" 
            logger.info("Activating Google Chrome window")
            self.dll.AU3_WinActivate(self.window_info, "")
            logger.info("Waiting Google Chrome window to be active")
            if not self.dll.AU3_WinWaitActive(self.window_info, "", 10):
                return "could not activate Google Chrome"
        else:
            return "Google Chrome window didn't exist"

        return None

    def browse(self, url = "www.python.org"):
        logger.info("Checking existing Google Chrome window")
        if self.dll.AU3_WinExists(self.window_info, ""):
            logger.info("Activating Google Chrome window")
            self.dll.AU3_WinActivate(self.window_info, "")
            logger.info("Waiting Google Chrome window to be active")
            if not self.dll.AU3_WinWaitActive(self.window_info, "", 10):
                return "could not activate Google Chrome window"
        else:
            return "Google Chrome window didn't exist"

        time.sleep(2)
        for letter in url:
            logger.info("Checking if Google Chrome window is active")
            if not self.dll.AU3_WinActive(self.window_info, ""):
                break

            logger.info("Sending letter to Google Chrome window")
            if not self.dll.AU3_Send(letter, 1):
                return f"could not send {letter} to Google Chrome"
            rand = random.uniform(0.05, 0.15)
            time.sleep(rand)
            
        time.sleep(1)
        logger.info("Sending enter to Google Chrome window")
        if not self.dll.AU3_Send("{ENTER}", 0):
            return "could not send Enter key"
            
        return None