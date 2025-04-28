import time
import random
import os
from configs.logger import logger
from activities.apps.native_apps.base import NativeApp

class Notepad(NativeApp):
    def __init__(self):
        super().__init__()
        self.window_info = "[CLASS:Notepad]"
        
    def _get_executable_path(self):
        path = os.path.join(os.environ['WINDIR'], 'System32', 'notepad.exe')
        if os.path.exists(path):
            return path
        return None
        
    def check_existing_window(self):
        logger.info("Checking existing Notepad window")
        if self.dll.AU3_WinExists(self.window_info, ""):
            logger.info("Activating Notepad window")
            self.dll.AU3_WinActivate(self.window_info, "")
            logger.info("Waiting Notepad window to be active")
            if not self.dll.AU3_WinWaitActive(self.window_info, "", 10):
                return "could not activate Notepad window"
        else:
            return "Notepad window didn't exist"
        
        return None
        
    def create_window(self):
        logger.info("Getting Notepad executable path")
        executable_path = self._get_executable_path()
        if not executable_path:
            return "could not get Notepad executable path"
        
        logger.info("Creating new Notepad window")
        if not self.dll.AU3_Run(executable_path, "", 1):
            return "could not run Notepad"

        time.sleep(2)
        logger.info("Checking existing Notepad window")
        if self.dll.AU3_WinExists(self.window_info, ""):
            logger.info("Getting Notepad window handle")
            handle = self.dll.AU3_WinGetHandle(self.window_info, "")
            if not handle:
                return "could not get window handle"
            else:
                self.window_info = "[HANDLE:%s]" % f"{handle:08X}" 
            logger.info("Activating Notepad window")
            self.dll.AU3_WinActivate(self.window_info, "")
            logger.info("Waiting Notepad window to be active")
            if not self.dll.AU3_WinWaitActive(self.window_info, "", 10):
                return "could not activate Notepad"
        else:
            return "Notepad window didn't exist"
        
        time.sleep(2)
        logger.info("Maximizing Notepad window")
        if not self.dll.AU3_WinSetState(self.window_info, "", 3):
            return "could not maximize Notepad window"
        
        return None
    
    def open_file(self, path):
        if not path:
            return "path must be provided"
        
        if not os.path.exists(path):
            return f"file path '{path}' does not exist"
        
        if not os.path.isfile(path):
            return f"path '{path}' is not a file"
        
        logger.info("Opening file with notepad")
        if not self.dll.AU3_Run(f'C:\\Windows\\System32\\notepad.exe "{path}"', "", 1):
            return "could not open file with notepad"
                
        time.sleep(2)
        logger.info("Checking existing Notepad window")
        if self.dll.AU3_WinExists(self.window_info, ""):
            logger.info("Getting Notepad window handle")
            handle = self.dll.AU3_WinGetHandle(self.window_info, "")
            if not handle:
                return "could not get window handle"
            else:
                self.window_info = "[HANDLE:%s]" % f"{handle:08X}" 
            logger.info("Activating Notepad window")
            self.dll.AU3_WinActivate(self.window_info, "")
            logger.info("Waiting Notepad window to be active")
            if not self.dll.AU3_WinWaitActive(self.window_info, "", 10):
                return "could not activate Notepad"
        else:
            return "Notepad window didn't exist"
        
        return None
    
    def write_file(self, text = ""):
        err = self.check_existing_window()
        if err:
            return err
        
        if text == "":
            text = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. In ultricies cursus sagittis."
            
        logger.info("Moving cursor to the bottom of the file")
        if not self.dll.AU3_Send("^{END}", 0):
            return "could not send keys to move cursor"
        
        time.sleep(2)
        for letter in text:
            logger.info("Checking if Notepad window is active")
            if not self.dll.AU3_WinActive(self.window_info, ""):
                return "Notepad window is inactive"

            logger.info("Sending letter to Notepad window")
            if not self.dll.AU3_Send(letter, 1):
                return f"could not send {letter} to Notepad"
            rand = random.uniform(0.05, 0.15)
            time.sleep(rand)
            
        return None
    
    def scroll(self, direction = "down", clicks = 10, scroll_delay = 0.05):
        if direction != "up" and direction != "down":
            return "invalid scroll direction"
        
        err = self.check_existing_window()
        if err:
            return err

        time.sleep(2)
        for _ in range(clicks):
            logger.info("Checking if Notepad window is active")
            if not self.dll.AU3_WinActive(self.window_info, ""):
                return "Notepad window is inactive"
            
            logger.info("Scrolling Mouse wheel")
            time.sleep(scroll_delay)
            if not self.dll.AU3_MouseWheel(direction, 1):
                return "could not scroll mouse wheel"
        
        return None
    
    def save_file(self):
        err = self.check_existing_window()
        if err:
            return err
        
        logger.info("Saving file")
        if not self.dll.AU3_Send("^s", 0):
            return "could not send ctrl+s to save file"
        
        return None