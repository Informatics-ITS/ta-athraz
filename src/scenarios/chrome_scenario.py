import time
import random
from utils.autoit_dll import AutoItDLL
from configs.logger import logger

class ChromeScenario:
    def __init__(self):
        self.dll = AutoItDLL().dll
        self.window_info = "[TITLE:New Tab - Google Chrome; CLASS:Chrome_WidgetWin_1]"
    
    def run(self):
        logger.info("Running Google Chrome")
        if not self.dll.AU3_Run("C:\\Program Files\\Google\Chrome\\Application\\chrome.exe", "", 1):
            return "could not run Google Chrome"
        
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
            logger.info("Waiting Microsoft Window to be active")
            if not self.dll.AU3_WinWaitActive(self.window_info, "", 10):
                return "could not activate Google Chrome"
        else:
            return "Google Chrome window didn't exist"

        return None
    
    def browse(self, text = ""):
        err = self.run()
        if err:
            return err
            
        if text == "":
            text = "ayam geprek"
            
        time.sleep(2)
        for letter in text:
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