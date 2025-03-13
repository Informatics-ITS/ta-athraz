import time
import random
from utils.autoit_dll import AutoItDLL
from configs.logger import logger

class WordScenario:
    def __init__(self):
        self.dll = AutoItDLL().dll
        self.window_info = "[CLASS:OpusApp]"
        
    def run(self):
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
            logger.info("Waiting Microsoft Window to be active")
            if not self.dll.AU3_WinWaitActive(self.window_info, "", 10):
                return "could not activate Microsoft Word"
        else:
            return "Microsoft Word window didn't exist"

        time.sleep(2)
        logger.info("Sending enter to Microsoft Word window")
        if not self.dll.AU3_Send("{ENTER}", 0):
            return "could not send Enter key"
        
        return None
    
    def write(self, text = ""):
        logger.info("Checking existing Microsoft Word window")
        if not self.dll.AU3_WinExists(self.window_info, ""):
            err = self.run()
            if err:
                return err
        else:    
            logger.info("Activating Microsoft Word window")
            self.dll.AU3_WinActivate(self.window_info, "")
            logger.info("Waiting Microsoft Word window to be active")
            if not self.dll.AU3_WinWaitActive(self.window_info, "", 10):
                return "could not activate Microsoft Word"
            
        if text == "":
            text = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. In ultricies cursus sagittis."

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