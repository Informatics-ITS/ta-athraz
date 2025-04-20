import time
import random
import os
from configs.logger import logger
from activities.files.base import File

class FileExplorer(File):
    def __init__(self):
        super().__init__()
        self.window_info = "[TITLE:Home - File Explorer; CLASS:CabinetWClass]"
        
    def _check_existing_window(self):
        logger.info("Checking existing File Explorer window")
        if self.dll.AU3_WinExists(self.window_info, ""):
            logger.info("Activating File Explorer window")
            self.dll.AU3_WinActivate(self.window_info, "")
            logger.info("Waiting File Explorer window to be active")
            if not self.dll.AU3_WinWaitActive(self.window_info, "", 10):
                return "could not activate File Explorer window"
        else:
            return "File Explorer window didn't exist"
        
        return None
    
    def create_window(self):
        logger.info("Running File Explorer")
        if not self.dll.AU3_Run("explorer.exe", "", 1):
            return "could not run File Explorer"

        time.sleep(2)
        logger.info("Checking existing File Explorer window")
        if self.dll.AU3_WinExists(self.window_info, ""):
            logger.info("Getting File Explorer window handle")
            handle = self.dll.AU3_WinGetHandle(self.window_info, "")
            if not handle:
                return "could not get window handle"
            else:
                self.window_info = "[HANDLE:%s]" % f"{handle:08X}" 
            logger.info("Activating File Explorer window")
            self.dll.AU3_WinActivate(self.window_info, "")
            logger.info("Waiting File Explorer window to be active")
            if not self.dll.AU3_WinWaitActive(self.window_info, "", 10):
                return "could not activate File Explorer"
        else:
            return "File Explorer window didn't exist"
        
        time.sleep(2)
        logger.info("Maximizing File Explorer window")
        if not self.dll.AU3_WinSetState(self.window_info, "", 3):
            return "could not maximize File Explorer window"
        
        return None
    
    def open_file(self, path):
        if not path:
            return "file path must be provided"
        
        if not os.path.exists(path):
            return f"file path '{path}' does not exist"
        
        if not os.path.isfile(path):
            return f"path '{path}' is not a file"
    
        err = self._check_existing_window()
        if err:
            return err
        
        time.sleep(1)
        logger.info("Opening File Explorer address bar")
        if not self.dll.AU3_Send("^l", 0):
            return "could not send keys to open address bar"
        
        time.sleep(1)
        for letter in path:
            logger.info("Checking if File Explorer window is active")
            if not self.dll.AU3_WinActive(self.window_info, ""):
                break

            logger.info("Sending letter to File Explorer window")
            if not self.dll.AU3_Send(letter, 1):
                return f"could not send {letter} to File Explorer"
            rand = random.uniform(0.05, 0.15)
            time.sleep(rand)
        
        time.sleep(1)
        logger.info("Sending Enter key to open file")
        if not self.dll.AU3_Send("{ENTER}", 0):
            return "could not send enter key to open file"
        
        return None
    
    def change_directory(self, path):
        if not path:
            return "directory path must be provided"
        
        if not os.path.exists(path):
            return f"directory '{path}' does not exist"
        
        if not os.path.isdir(path):
            return f"path '{path}' is not a directory"
        
        err = self._check_existing_window()
        if err:
            return err
        
        time.sleep(1)
        logger.info("Opening File Explorer address bar")
        if not self.dll.AU3_Send("^l", 0):
            return "could not send keys to open address bar"
        
        time.sleep(1)
        for letter in path:
            logger.info("Checking if File Explorer window is active")
            if not self.dll.AU3_WinActive(self.window_info, ""):
                break

            logger.info("Sending letter to File Explorer window")
            if not self.dll.AU3_Send(letter, 1):
                return f"could not send {letter} to File Explorer"
            rand = random.uniform(0.05, 0.15)
            time.sleep(rand)
        
        time.sleep(1)
        logger.info("Sending Enter key to change directory")
        if not self.dll.AU3_Send("{ENTER}", 0):
            return "could not send enter key to change directory"
        
        return None
    
    def create_directory(self, parent_path, dir_name):
        if not parent_path or not dir_name:
            return "parent path and directory name must be provided"
        
        full_path = os.path.join(parent_path, dir_name)
        if os.path.exists(full_path):
            return f"directory '{dir_name}' already exist at '{parent_path}'"
        
        err = self.change_directory(parent_path)
        if err:
            return err
        
        time.sleep(1)
        logger.info("Sending keys to access new directory shortcut")
        if not self.dll.AU3_Send("^+n", 0):
            return "could not send keys to access new directory shortcut"
        
        time.sleep(1)
        for letter in dir_name:
            logger.info("Checking if File Explorer window is active")
            if not self.dll.AU3_WinActive(self.window_info, ""):
                break

            logger.info("Sending letter to File Explorer window")
            if not self.dll.AU3_Send(letter, 1):
                return f"could not send {letter} to File Explorer"
            rand = random.uniform(0.05, 0.15)
            time.sleep(rand)
            
        time.sleep(1)
        logger.info("Sending Enter key to create new directory")
        if not self.dll.AU3_Send("{ENTER}", 0):
            return "could not send enter key to create new directory"
        
        return None