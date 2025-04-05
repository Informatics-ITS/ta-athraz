import time
import random
from configs.logger import logger
from scenarios.base.browser_scenario import BrowserScenario

class ChromeScenario(BrowserScenario):
    def __init__(self):
        super().__init__()
        self.window_info = "[TITLE:New Tab - Google Chrome; CLASS:Chrome_WidgetWin_1]"

    def run(self):
        logger.info("Running Google Chrome")
        if not self.dll.AU3_Run("C:\\Program Files\\Google\Chrome\\Application\\chrome.exe", "", 1):
            return "could not run Google Chrome"
        
    def browse(self, url = ""):
        err = self.run()
        if err:
            return err
            
        if url == "":
            url = "www.python.org"
            
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