import time
import random
from configs.logger import logger
from activities.apps.browsers.base import Browser

class MozillaFirefox(Browser):
    def __init__(self):
        super().__init__()
        self.window_info = "[CLASS:MozillaWindowClass]"
        
    def _check_existing_window(self):
        logger.info("Checking existing Mozilla Firefox window")
        if self.dll.AU3_WinExists(self.window_info, ""):
            logger.info("Activating Mozilla Firefox window")
            self.dll.AU3_WinActivate(self.window_info, "")
            logger.info("Waiting Mozilla Firefox window to be active")
            if not self.dll.AU3_WinWaitActive(self.window_info, "", 10):
                return "could not activate Mozilla Firefox window"
        else:
            return "Mozilla Firefox window didn't exist"

        return None
    
    def create_window(self):
        logger.info("Creating new Mozilla Firefox window")
        if not self.dll.AU3_Run("C:\\Users\\ASUS\\AppData\\Local\\Mozilla Firefox\\firefox.exe", "", 1):
            return "could not run firefox.exe"
        
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
    
    def create_tab(self):
        err = self._check_existing_window()
        if err:
            return err

        logger.info("Creating new Mozilla Firefox tab")
        if not self.dll.AU3_Send("^t", 0):
            return "Could not send keys to create new tab"

        return None

    
    def browse(self, url = "www.python.org"):
        err = self._check_existing_window()
        if err:
            return err

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
    
    def scroll(self, direction = "down", clicks = 10, scroll_delay = 0.05):
        if direction != "up" and direction != "down":
            return "invalid scroll direction"
        
        err = self._check_existing_window()
        if err:
            return err

        time.sleep(2)
        for _ in range(clicks):
            logger.info("Checking if Mozilla Firefox window is active")
            if not self.dll.AU3_WinActive(self.window_info, ""):
                break
            
            logger.info("Scrolling Mouse wheel")
            time.sleep(scroll_delay)
            if not self.dll.AU3_MouseWheel(direction, 1):
                return "Cannot scroll mouse wheel"
        
        return None
    
    def zoom_in(self, count = 1):
        err = self._check_existing_window()
        if err:
            return err

        time.sleep(2)
        for _ in range(count):
            logger.info("Sending keys to zoom in Mozilla Firefox")
            if not self.dll.AU3_Send("^{+}", 0):
                return "Could not send keys to zoom in"
            time.sleep(0.2)

        return None

        
    def zoom_out(self, count = 1):
        err = self._check_existing_window()
        if err:
            return err

        time.sleep(2)
        for _ in range(count):
            logger.info("Sending keys to zoom out Mozilla Firefox")
            if not self.dll.AU3_Send("^{-}", 0):
                return "Could not send keys to zoom out"
            time.sleep(0.2)

        return None