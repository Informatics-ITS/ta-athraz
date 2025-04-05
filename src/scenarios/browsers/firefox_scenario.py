import time
import random
from configs.logger import logger
from scenarios.base.browser_scenario import BrowserScenario

class FirefoxScenario(BrowserScenario):
    def __init__(self):
        super().__init__()
        self.window_info = "[CLASS:MozillaWindowClass]"
    
    def run(self):
        logger.info("Running Mozilla Firefox")
        if not self.dll.AU3_Run("C:\\Users\\ASUS\\AppData\\Local\\Mozilla Firefox\\firefox.exe", "", 1):
            return "could not run Mozilla Firefox"
        
        time.sleep(2)
        logger.info("Checking existing Mozilla Firefox window")
        if self.dll.AU3_WinExists(self.window_info, ""):
            logger.info("Getting Mozilla Firefox window handle")
            handle = self.dll.AU3_WinGetHandle(self.window_info, "")
            if not handle:
                return "could not get window handle"
            else:
                self.window_info = "[HANDLE:%s]" % f"{handle:08X}" 
            logger.info("Activating Mozilla Firefox window")
            self.dll.AU3_WinActivate(self.window_info, "")
            logger.info("Waiting Mozilla Firefox window to be active")
            if not self.dll.AU3_WinWaitActive(self.window_info, "", 10):
                return "could not activate Mozilla Firefox"
        else:
            return "Mozilla Firefox window didn't exist"

        return None
    
    def browse(self, url = ""):
        logger.info("Checking existing Mozilla Firefox window")
        if not self.dll.AU3_WinExists(self.window_info, ""):
            err = self.run()
            if err:
                return err
        else:    
            logger.info("Activating Mozilla Firefox window")
            self.dll.AU3_WinActivate(self.window_info, "")
            logger.info("Waiting Mozilla Firefox window to be active")
            if not self.dll.AU3_WinWaitActive(self.window_info, "", 10):
                return "could not activate Mozilla Firefox"
            logger.info("Createing new Mozilla Firefox tab")
            self.dll.AU3_Send("^t", 0)
            
        if url == "":
            url = "www.python.org"
            
        time.sleep(2)
        for letter in url:
            logger.info("Checking if Mozilla Firefox window is active")
            if not self.dll.AU3_WinActive(self.window_info, ""):
                break

            logger.info("Sending letter to Mozilla Firefox window")
            if not self.dll.AU3_Send(letter, 1):
                return f"could not send {letter} to Mozilla Firefox"
            rand = random.uniform(0.05, 0.15)
            time.sleep(rand)
            
        time.sleep(1)
        logger.info("Sending enter to Mozilla Firefox window")
        if not self.dll.AU3_Send("{ENTER}", 0):
            return "could not send Enter key"
            
        return None