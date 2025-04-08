import time
import random
import os
from configs.logger import logger
from activities.apps.native_apps.base import NativeApp

class MicrosoftWord(NativeApp):
    def __init__(self):
        super().__init__()
        self.window_info = "[CLASS:OpusApp]"
        
    def _check_existing_window(self):
        logger.info("Checking existing Microsoft Word window")
        if self.dll.AU3_WinExists(self.window_info, ""):
            logger.info("Activating Microsoft Word window")
            self.dll.AU3_WinActivate(self.window_info, "")
            logger.info("Waiting Microsoft Word window to be active")
            if not self.dll.AU3_WinWaitActive(self.window_info, "", 10):
                return "could not activate Microsoft Word window"
        else:
            return "Microsoft Word window didn't exist"
        
        return None
        
    def create(self):
        logger.info("Running Microsoft Word")
        if not self.dll.AU3_Run("C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE", "", 1):
            return "could not run Microsoft Word"

        time.sleep(2)
        logger.info("Checking existing Microsoft Word window")
        if self.dll.AU3_WinExists(self.window_info, ""):
            logger.info("Getting Microsoft Word window handle")
            handle = self.dll.AU3_WinGetHandle(self.window_info, "")
            if not handle:
                return "could not get window handle"
            else:
                self.window_info = "[HANDLE:%s]" % f"{handle:08X}" 
            logger.info("Activating Microsoft Word window")
            self.dll.AU3_WinActivate(self.window_info, "")
            logger.info("Waiting Microsoft Word window to be active")
            if not self.dll.AU3_WinWaitActive(self.window_info, "", 10):
                return "could not activate Microsoft Word"
        else:
            return "Microsoft Word window didn't exist"

        time.sleep(2)
        logger.info("Sending enter to Microsoft Word window")
        if not self.dll.AU3_Send("{ENTER}", 0):
            return "could not send Enter key to create new docx"
        
        return None
        
    def open(self, path):
        if not path:
            return "path must be provided"
        
        if not os.path.exists(path):
            return f"file path '{path}' does not exist"
        
        if not os.path.isfile(path):
            return f"path '{path}' is not a file"
        
        logger.info("Opening docx file")
        if not self.dll.AU3_Run(f'C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE "{path}"', "", 1):
            return "could not open docx file"

        time.sleep(2)
        logger.info("Checking existing Microsoft Word window")
        if self.dll.AU3_WinExists(self.window_info, ""):
            logger.info("Getting Microsoft Word window handle")
            handle = self.dll.AU3_WinGetHandle(self.window_info, "")
            if not handle:
                return "could not get window handle"
            else:
                self.window_info = "[HANDLE:%s]" % f"{handle:08X}" 
            logger.info("Activating Microsoft Word window")
            self.dll.AU3_WinActivate(self.window_info, "")
            logger.info("Waiting Microsoft Word window to be active")
            if not self.dll.AU3_WinWaitActive(self.window_info, "", 10):
                return "could not activate Microsoft Word"
        else:
            return "Microsoft Word window didn't exist"

        return None
    
    def write(self, text = ""):
        err = self._check_existing_window()
        if err:
            return err
            
        if text == "":
            text = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. In ultricies cursus sagittis."
            
        logger.info("Moving cursor to the bottom of docx file")
        if not self.dll.AU3_Send("^{END}", 0):
            return "could not send keys to move cursor"

        time.sleep(2)
        for letter in text:
            logger.info("Checking if Microsoft Word window is active")
            if not self.dll.AU3_WinActive(self.window_info, ""):
                break

            logger.info("Sending letter to Microsoft Word window")
            if not self.dll.AU3_Send(letter, 1):
                return f"could not send {letter} to Microsoft Word"
            rand = random.uniform(0.05, 0.15)
            time.sleep(rand)
            
        return None
    
    def scroll(self, direction = "down", clicks = 10, scroll_delay = 0.05):
        if direction != "up" and direction != "down":
            return "invalid scroll direction"
        
        err = self._check_existing_window()
        if err:
            return err

        time.sleep(2)
        for _ in range(clicks):
            logger.info("Checking if Microsoft Word window is active")
            if not self.dll.AU3_WinActive(self.window_info, ""):
                break
            
            logger.info("Scrolling Mouse wheel")
            time.sleep(scroll_delay)
            if not self.dll.AU3_MouseWheel(direction, 1):
                return "Cannot scroll mouse wheel"
        
        return None
    
    def save(self):
        err = self._check_existing_window()
        if err:
            return err
        
        logger.info("Saving docx file")
        if not self.dll.AU3_Send("^s", 0):
            return "could not send ctrl+s to save docx file"
        
        return None