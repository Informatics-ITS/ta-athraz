import time
import random
import os
from configs.logger import logger
from activities.apps.native_apps.base import NativeApp

class MicrosoftExcel(NativeApp):
    def __init__(self):
        super().__init__()
        self.window_info = "[CLASS:XLMAIN]"
        self.goto_window_info = "[TITLE:Go To; CLASS:bosa_sdm_XL9]"
        
    def check_existing_window(self):
        logger.info("Checking existing Microsoft Excel window")
        if self.dll.AU3_WinExists(self.window_info, ""):
            logger.info("Activating Microsoft Excel window")
            self.dll.AU3_WinActivate(self.window_info, "")
            logger.info("Waiting Microsoft Excel window to be active")
            if not self.dll.AU3_WinWaitActive(self.window_info, "", 10):
                return "could not activate Microsoft Excel window"
        else:
            return "Microsoft Excel window didn't exist"
        
        return None
    
    def create_window(self):
        logger.info("Running Microsoft Excel")
        if not self.dll.AU3_Run("C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE --start-maximized", "", 1):
            return "could not run Microsoft Excel"

        time.sleep(2)
        logger.info("Checking existing Microsoft Excel window")
        if self.dll.AU3_WinExists(self.window_info, ""):
            logger.info("Getting Microsoft Excel window handle")
            handle = self.dll.AU3_WinGetHandle(self.window_info, "")
            if not handle:
                return "could not get window handle"
            else:
                self.window_info = "[HANDLE:%s]" % f"{handle:08X}" 
            logger.info("Activating Microsoft Excel window")
            self.dll.AU3_WinActivate(self.window_info, "")
            logger.info("Waiting Microsoft Excel window to be active")
            if not self.dll.AU3_WinWaitActive(self.window_info, "", 10):
                return "could not activate Microsoft Excel"
        else:
            return "Microsoft Excel window didn't exist"
        
        time.sleep(2)
        logger.info("Maximizing Microsoft Excel window")
        if not self.dll.AU3_WinSetState(self.window_info, "", 3):
            return "could not maximize Microsoft Excel window"

        time.sleep(2)
        logger.info("Sending enter to Microsoft Excel window")
        if not self.dll.AU3_Send("{ENTER}", 0):
            return "could not send Enter key to create new xlsx"
        
        return None
        
    def open_xlsx(self, path):
        if not path:
            return "path must be provided"
        
        if not os.path.exists(path):
            return f"file path '{path}' does not exist"
        
        if not os.path.isfile(path):
            return f"path '{path}' is not a file"
        
        logger.info("Opening xlsx file")
        if not self.dll.AU3_Run(f'C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE "{path}"', "", 1):
            return "could not open xlsx file"

        time.sleep(2)
        logger.info("Checking existing Microsoft Excel window")
        if self.dll.AU3_WinExists(self.window_info, ""):
            logger.info("Getting Microsoft Excel window handle")
            handle = self.dll.AU3_WinGetHandle(self.window_info, "")
            if not handle:
                return "could not get window handle"
            else:
                self.window_info = "[HANDLE:%s]" % f"{handle:08X}" 
            logger.info("Activating Microsoft Excel window")
            self.dll.AU3_WinActivate(self.window_info, "")
            logger.info("Waiting Microsoft Excel window to be active")
            if not self.dll.AU3_WinWaitActive(self.window_info, "", 10):
                return "could not activate Microsoft Excel"
        else:
            return "Microsoft Excel window didn't exist"
        
        time.sleep(2)
        logger.info("Maximizing Microsoft Excel window")
        if not self.dll.AU3_WinSetState(self.window_info, "", 3):
            return "could not maximize Microsoft Excel window"

        return None
        
    def change_cell(self, target_cell = "A1"):
        err = self.check_existing_window()
        if err:
            return err
        
        time.sleep(2)
        logger.info("Sending keys to open Go To window")
        if not self.dll.AU3_Send("^g", 0):
            return "could not send keys to open Go To window"
        
        time.sleep(1)
        logger.info("Checking existing Go To window")
        if self.dll.AU3_WinExists(self.goto_window_info, ""):
            logger.info("Activating Go To window")
            self.dll.AU3_WinActivate(self.goto_window_info, "")
            logger.info("Waiting Go To window to be active")
            if not self.dll.AU3_WinWaitActive(self.goto_window_info, "", 10):
                return "could not activate Go To window"
        else:
            return "Go To window didn't exist"
        
        time.sleep(1)
        logger.info("Sending target cell to Go To window")
        if not self.dll.AU3_Send(target_cell, 1):
            return "could not send target cell to Go To window"
        
        time.sleep(1)
        logger.info("Sending enter to Go To window")
        if not self.dll.AU3_Send("{ENTER}", 0):
            return "could not send Enter key to change cell"
        
        return None

    def create_sheet(self):
        err = self.check_existing_window()
        if err:
            return err
        
        time.sleep(2)
        logger.info("Sending keys to create new sheet")
        if not self.dll.AU3_Send("+{F11}", 0):
            return "could not send keys to create new sheet"
        
        return None
        
    def rename_sheet(self, new_name):
        if new_name == "":
            return "new name must be provided"
        
        err = self.check_existing_window()
        if err:
            return err
        
        time.sleep(2)
        logger.info("Sending keys to create new sheet")
        if not self.dll.AU3_Send("!hor", 0):
            return "could not send keys to create new sheet"
        
        time.sleep(2)
        for letter in new_name:
            logger.info("Checking if Microsoft Excel window is active")
            if not self.dll.AU3_WinActive(self.window_info, ""):
                return "Microsoft Excel window is inactive"

            logger.info("Sending letter to Microsoft Excel window")
            if not self.dll.AU3_Send(letter, 1):
                return f"could not send {letter} to Microsoft Excel"
            rand = random.uniform(0.05, 0.15)
            time.sleep(rand)
            
        time.sleep(1)
        logger.info("Sending enter to save new name")
        if not self.dll.AU3_Send("{ENTER}", 0):
            return "could not send Enter key to save new name"
        
        return None
        
    def change_sheet(self, direction = "right", count = 1):
        if direction != "left" and direction != "right":
            return "invalid change sheet direction"
        
        err = self.check_existing_window()
        if err:
            return err
        
        time.sleep(2)
        for _ in range(count):
            logger.info("Checking if Microsoft Excel window is active")
            if not self.dll.AU3_WinActive(self.window_info, ""):
                return "Microsoft Excel window is inactive"
            
            logger.info("Sending keys to change sheets")
            if direction == "left":
                if not self.dll.AU3_Send("^{PGUP}", 0):
                    return "could not send keys to change sheets"
            else:
                if not self.dll.AU3_Send("^{PGDN}", 0):
                    return "could not send keys to change sheets"
            rand = random.uniform(0.05, 0.15)
            time.sleep(rand)
        
        return None
        
    def write_cell(self, text = ""):
        err = self.check_existing_window()
        if err:
            return err
            
        if text == "":
            text = "Lorem ipsum dolor sit amet."

        time.sleep(2)
        for letter in text:
            logger.info("Checking if Microsoft Excel window is active")
            if not self.dll.AU3_WinActive(self.window_info, ""):
                return "Microsoft Excel window is inactive"

            logger.info("Sending letter to Microsoft Excel window")
            if not self.dll.AU3_Send(letter, 1):
                return f"could not send {letter} to Microsoft Excel"
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
            logger.info("Checking if Microsoft Excel window is active")
            if not self.dll.AU3_WinActive(self.window_info, ""):
                return "Microsoft Excel window is inactive"
            
            logger.info("Scrolling Mouse wheel")
            time.sleep(scroll_delay)
            if not self.dll.AU3_MouseWheel(direction, 1):
                return "could not scroll mouse wheel"
        
        return None
    
    def save_xlsx(self):
        err = self.check_existing_window()
        if err:
            return err
        
        logger.info("Saving xlsx file")
        if not self.dll.AU3_Send("^s", 0):
            return "could not send ctrl+s to save xlsx file"
        
        return None