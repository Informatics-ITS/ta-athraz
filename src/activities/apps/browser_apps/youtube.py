import time
import random
from configs.logger import logger
from activities.apps.browser_apps.base import BrowserApp

class YouTube(BrowserApp):
    def open(self):
        logger.info("Checking browser window")
        err = self.browser.check_existing_window()
        if err:
            logger.info("Creating new browser window")
            err = self.browser.create_window()
            if err:
                return err
        else:
            logger.info("Creating new browser tab")
            err = self.browser.create_tab()
            if err:
                return err
        
        logger.info("Browsing YouTube")
        err = self.browser.browse("https://youtube.com")
        if err:
            return err
        
        return None
        
    def search(self, text="python tutorial"):
        logger.info("Checking browser window")
        err = self.browser.check_existing_window()
        if err:
            return err
        
        time.sleep(2)
        logger.info("Sending keys to go to search box")
        if not self.dll.AU3_Send("/", 0):
            return "could not send keys to go to search box"
        
        time.sleep(1)
        logger.info("Sending keys to clear search box")
        if not self.dll.AU3_Send("^a{BACKSPACE}", 0):
            return "could not send keys to clear search box"
        
        time.sleep(2)
        for letter in text:
            logger.info("Checking if browser window is active")
            if not self.dll.AU3_WinActive(self.browser.window_info, ""):
                return "browser window is inactive while filling YouTube search box"

            logger.info("Sending text to YouTube search box")
            if not self.dll.AU3_Send(letter, 1):
                return f"could not send {letter} to YouTube search box"
            rand = random.uniform(0.05, 0.15)
            time.sleep(rand)
            
        time.sleep(1)
        if not self.dll.AU3_Send("{ENTER}", 0):
            return "could not send enter key to search"
        
        return None
    
    def toggle_pause(self):
        logger.info("Checking browser window")
        err = self.browser.check_existing_window()
        if err:
            return err
        
        time.sleep(1)
        logger.info("Sending keys to toggle pause")
        if not self.dll.AU3_Send("{SPACE}", 0):
            return "could not send keys to toggle pause"
        
        return None
    
    def toggle_mute(self):
        logger.info("Checking browser window")
        err = self.browser.check_existing_window()
        if err:
            return err
        
        time.sleep(1)
        logger.info("Sending keys to toggle mute")
        if not self.dll.AU3_Send("m", 0):
            return "could not send keys to toggle mute"
        
        return None
    
    def toggle_subtitle(self):
        logger.info("Checking browser window")
        err = self.browser.check_existing_window()
        if err:
            return err
        
        time.sleep(1)
        logger.info("Sending keys to toggle subtitle")
        if not self.dll.AU3_Send("c", 0):
            return "could not send keys to toggle subtitle"
        
        return None
    
    def toggle_fullscreen(self):
        logger.info("Checking browser window")
        err = self.browser.check_existing_window()
        if err:
            return err
        
        time.sleep(1)
        logger.info("Sending keys to toggle fullscreen")
        if not self.dll.AU3_Send("f", 0):
            return "could not send keys to toggle fullscreen"
        
        return None
        
    def toggle_cinema_mode(self):
        logger.info("Checking browser window")
        err = self.browser.check_existing_window()
        if err:
            return err
        
        time.sleep(1)
        logger.info("Sending keys to toggle cinema mode")
        if not self.dll.AU3_Send("t", 0):
            return "could not send keys to toggle cinema mode"
        
        return None
        
    def next_video(self, count = 1):
        logger.info("Checking browser window")
        err = self.browser.check_existing_window()
        if err:
            return err
        
        time.sleep(1)
        for _ in range(count):
            logger.info("Checking if browser window is active")
            if not self.dll.AU3_WinActive(self.browser.window_info, ""):
                return "browser window is inactive"
            
            logger.info("Sending keys to move to next video")
            if not self.dll.AU3_Send("+n", 0):
                return "could not send keys to move to next video"
            time.sleep(0.2)
        
        return None
        
    def previous_video(self, count = 1):
        logger.info("Checking browser window")
        err = self.browser.check_existing_window()
        if err:
            return err
        
        time.sleep(1)
        for _ in range(count):
            logger.info("Checking if browser window is active")
            if not self.dll.AU3_WinActive(self.browser.window_info, ""):
                return "browser window is inactive"
            
            logger.info("Sending keys to move to previous video")
            if not self.dll.AU3_Send("+p", 0):
                return "could not send keys to move to previous video"
            time.sleep(0.2)
        
        return None