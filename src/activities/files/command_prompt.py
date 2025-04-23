import time
import random
import os
from configs.logger import logger
from activities.files.base import File

class CommandPrompt(File):
    def __init__(self):
        super().__init__()
        self.window_info = "[TITLE:C:\Windows\system32\cmd.exe; CLASS:CASCADIA_HOSTING_WINDOW_CLASS]"
    
    def check_existing_window(self):
        logger.info("Checking existing Command Prompt window")
        if self.dll.AU3_WinExists(self.window_info, ""):
            logger.info("Activating Command Prompt window")
            self.dll.AU3_WinActivate(self.window_info, "")
            logger.info("Waiting Command Prompt window to be active")
            if not self.dll.AU3_WinWaitActive(self.window_info, "", 10):
                return "could not activate Command Prompt window"
        else:
            return "Command Prompt window didn't exist"
        
        return None
    
    def create_window(self):
        logger.info("Running Command Prompt")
        if not self.dll.AU3_Run("cmd.exe /c start", "", 1):
            return "could not run Command Prompt"

        time.sleep(2)
        logger.info("Checking existing Command Prompt window")
        if self.dll.AU3_WinExists(self.window_info, ""):
            logger.info("Getting Command Prompt window handle")
            handle = self.dll.AU3_WinGetHandle(self.window_info, "")
            if not handle:
                return "could not get window handle"
            else:
                self.window_info = "[HANDLE:%s]" % f"{handle:08X}" 
            logger.info("Activating Command Prompt window")
            self.dll.AU3_WinActivate(self.window_info, "")
            logger.info("Waiting Command Prompt window to be active")
            if not self.dll.AU3_WinWaitActive(self.window_info, "", 10):
                return "could not activate Command Prompt"
        else:
            return "Command Prompt window didn't exist"
        
        time.sleep(2)
        logger.info("Maximizing Command Prompt window")
        if not self.dll.AU3_WinSetState(self.window_info, "", 3):
            return "could not maximize Command Prompt window"
        
        return None

    def open_file(self, path):
        if not path:
            return "file path must be provided"
        
        if not os.path.exists(path):
            return f"file path '{path}' does not exist"
        
        if not os.path.isfile(path):
            return f"path '{path}' is not a file"
        
        err = self.check_existing_window()
        if err:
            return err
        
        time.sleep(2)
        cmdline = f'start "" "{path}"'
        for letter in cmdline:
            logger.info("Checking if Command Prompt window is active")
            if not self.dll.AU3_WinActive(self.window_info, ""):
                break

            logger.info("Sending letter to Command Prompt window")
            if not self.dll.AU3_Send(letter, 1):
                return f"could not send {letter} to Command Prompt"
            rand = random.uniform(0.05, 0.15)
            time.sleep(rand)
        
        time.sleep(1)
        logger.info("Sending Enter key to execute open file command")
        if not self.dll.AU3_Send("{ENTER}", 0):
            return "could not send enter key to execute command"
        
        return None
    
    def change_directory(self, path):
        if not path:
            return "directory path must be provided"
        
        if not os.path.exists(path):
            return f"directory '{path}' does not exist"
        
        if not os.path.isdir(path):
            return f"path '{path}' is not a directory"
        
        err = self.check_existing_window()
        if err:
            return err
        
        time.sleep(2)
        cmdline = f'cd "{path}"'
        for letter in cmdline:
            logger.info("Checking if Command Prompt window is active")
            if not self.dll.AU3_WinActive(self.window_info, ""):
                break

            logger.info("Sending letter to Command Prompt window")
            if not self.dll.AU3_Send(letter, 1):
                return f"could not send {letter} to Command Prompt"
            rand = random.uniform(0.05, 0.15)
            time.sleep(rand)
        
        time.sleep(1)
        logger.info("Sending Enter key to execute change directory command")
        if not self.dll.AU3_Send("{ENTER}", 0):
            return "could not send enter key to execute command"
        
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
        
        time.sleep(2)
        cmdline = f'mkdir "{dir_name}"'
        for letter in cmdline:
            logger.info("Checking if Command Prompt window is active")
            if not self.dll.AU3_WinActive(self.window_info, ""):
                break

            logger.info("Sending letter to Command Prompt window")
            if not self.dll.AU3_Send(letter, 1):
                return f"could not send {letter} to Command Prompt"
            rand = random.uniform(0.05, 0.15)
            time.sleep(rand)
        
        time.sleep(1)
        logger.info("Sending Enter key to execute create directory command")
        if not self.dll.AU3_Send("{ENTER}", 0):
            return "could not send enter key to execute command"
        
        return None