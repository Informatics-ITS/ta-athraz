import time
import random
from configs.logger import logger
from activities.apps.browser_apps.base import BrowserApp

class Gmail(BrowserApp):
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
        
        logger.info("Browsing Gmail")
        err = self.browser.browse("https://mail.google.com")
        if err:
            return err
        
        return None
    
    def open_email(self):
        logger.info("Checking browser window")
        err = self.browser.check_existing_window()
        if err:
            return err
        
        time.sleep(1)
        logger.info("Sending enter key to open email")
        if not self.dll.AU3_Send("{ENTER}", 0):
            return "could not send enter key to open email"
        
        return None
        
    def next_email(self, count = 1):
        logger.info("Checking browser window")
        err = self.browser.check_existing_window()
        if err:
            return err
        
        time.sleep(1)
        for _ in range(count):
            logger.info("Sending key to go to next email")
            if not self.dll.AU3_Send("j", 0):
                return "could not send key to go to next email"
            time.sleep(0.2)
        
        return None
    
    def previous_email(self, count = 1):
        logger.info("Checking browser window")
        err = self.browser.check_existing_window()
        if err:
            return err
        
        time.sleep(1)
        for _ in range(count):
            logger.info("Sending key to go to previous email")
            if not self.dll.AU3_Send("k", 0):
                return "could not send key to go to previous email"
            time.sleep(0.2)
        
        return None
    
    def compose_email(self, recipients, subject, body):
        logger.info("Checking browser window")
        err = self.browser.check_existing_window()
        if err:
            return err
        
        time.sleep(1)
        logger.info("Sending key to open compose tab")
        if not self.dll.AU3_Send("c", 0):
            return "could not send key to open compose tab"
        
        time.sleep(2)
        logger.info("Filling recipients")
        for recipient in recipients:
            logger.info(f"Filling recipient {recipient}")
            for letter in recipient:
                logger.info("Checking if browser window is active")
                if not self.dll.AU3_WinActive(self.browser.window_info, ""):
                    return "browser window is inactive while filling email recipient"

                logger.info("Sending text to recipient field")
                if not self.dll.AU3_Send(letter, 1):
                    return f"could not send {letter} to recipient field"
                rand = random.uniform(0.05, 0.15)
                time.sleep(rand)
                
            time.sleep(1)
            if not self.dll.AU3_Send("{ENTER}", 0):
                return "could not send enter key to confirm recipient"
            
        time.sleep(2)
        logger.info("Sending key to go to subject field")
        if not self.dll.AU3_Send("{TAB}", 0):
            return "could not send key to go to subject field"
        
        time.sleep(2)
        for letter in subject:
            logger.info("Checking if browser window is active")
            if not self.dll.AU3_WinActive(self.browser.window_info, ""):
                return "browser window is inactive while filling email subject"

            logger.info("Sending text to subject field")
            if not self.dll.AU3_Send(letter, 1):
                return f"could not send {letter} to subject field"
            rand = random.uniform(0.05, 0.15)
            time.sleep(rand)
        
        time.sleep(2)
        logger.info("Sending key to go to body field")
        if not self.dll.AU3_Send("{TAB}", 0):
            return "could not send key to go to body field"
        
        time.sleep(2)
        for letter in body:
            logger.info("Checking if browser window is active")
            if not self.dll.AU3_WinActive(self.browser.window_info, ""):
                return "browser window is inactive while filling email body"

            logger.info("Sending text to body field")
            if not self.dll.AU3_Send(letter, 1):
                return f"could not send {letter} to body field"
            rand = random.uniform(0.05, 0.15)
            time.sleep(rand)
            
        time.sleep(2)
        logger.info("Sending keys to send email")
        if not self.dll.AU3_Send("^{ENTER}", 0):
            return "could not send keys to send email"
        
        return None
    
    def go_to_inbox(self):
        logger.info("Checking browser window")
        err = self.browser.check_existing_window()
        if err:
            return err
        
        time.sleep(1)
        logger.info("Sending keys to go to inbox")
        if not self.dll.AU3_Send("gi", 0):
            return "could not send keys to go to inbox"
        
        return None
    
    def go_to_drafts(self):
        logger.info("Checking browser window")
        err = self.browser.check_existing_window()
        if err:
            return err
        
        time.sleep(1)
        logger.info("Sending keys to go to inbox")
        if not self.dll.AU3_Send("gd", 0):
            return "could not send keys to go to inbox"
        
        return None
    
    def go_to_sent_messages(self):
        logger.info("Checking browser window")
        err = self.browser.check_existing_window()
        if err:
            return err
        
        time.sleep(1)
        logger.info("Sending keys to go to inbox")
        if not self.dll.AU3_Send("gt", 0):
            return "could not send keys to go to inbox"
        
        return None
    
    def select_all(self):
        logger.info("Checking browser window")
        err = self.browser.check_existing_window()
        if err:
            return err
        
        time.sleep(1)
        logger.info("Sending keys to select all email")
        if not self.dll.AU3_Send("*a", 0):
            return "could not send keys to select all email"
        
        return None
        
    def deselect_all(self):
        logger.info("Checking browser window")
        err = self.browser.check_existing_window()
        if err:
            return err
        
        time.sleep(1)
        logger.info("Sending keys to deselect all email")
        if not self.dll.AU3_Send("*n", 0):
            return "could not send keys to deselect all email"
        
        return None
    
    def select_unread(self):
        logger.info("Checking browser window")
        err = self.browser.check_existing_window()
        if err:
            return err
        
        time.sleep(1)
        logger.info("Sending keys to select unread email")
        if not self.dll.AU3_Send("*u", 0):
            return "could not send keys to select unread email"
        
        return None
    
    def mark_as_unread(self):
        logger.info("Checking browser window")
        err = self.browser.check_existing_window()
        if err:
            return err
        
        time.sleep(1)
        logger.info("Sending keys to mark selected email as unread")
        if not self.dll.AU3_Send("+u", 0):
            return "could not send keys to mark selected email as unread"
        
        return None
    
    def mark_as_read(self):
        logger.info("Checking browser window")
        err = self.browser.check_existing_window()
        if err:
            return err
        
        time.sleep(1)
        logger.info("Sending keys to mark selected email as read")
        if not self.dll.AU3_Send("+i", 0):
            return "could not send keys to mark selected email as read"
        
        return None
    
    def archive_email(self):
        logger.info("Checking browser window")
        err = self.browser.check_existing_window()
        if err:
            return err
        
        time.sleep(1)
        logger.info("Sending key to archive email")
        if not self.dll.AU3_Send("e", 0):
            return "could not send key to archive email"
        
        return None