import time
import random
import pyperclip
from configs.logger import logger
from activities.apps.browser_apps.base import BrowserApp

class GoogleForms(BrowserApp):        
    def fill_form(self, url, answers):
        if not url.startswith("https://forms.gle/") and not url.startswith("https://https://docs.google.com/forms/"):
            return "Invalid Google Forms URL"
        
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
        
        logger.info("Browsing Google Forms URL")
        err = self.browser.browse(url)
        if err:
            return err
        
        time.sleep(2)
        first = True
        for answer in answers:
            logger.info("Checking if browser window is active")
            if not self.dll.AU3_WinActive(self.browser.window_info, ""):
                return "browser window is inactive while moving to Google Forms questions"
            
            if first:
                logger.info("Moving to Google Forms question")
                if not self.dll.AU3_Send("{TAB}{TAB}{TAB}", 0):
                    return "could not send tab key to move to Google Forms question"
                first = False
            else:
                logger.info("Moving to next question")
                if not self.dll.AU3_Send("{TAB}", 0):
                    return "could not send tab key to move to next question"

            time.sleep(random.uniform(1, 3))
            
            logger.info("Sending answer to Google Forms")
            for letter in answer:
                logger.info("Checking if browser window is active")
                if not self.dll.AU3_WinActive(self.browser.window_info, ""):
                    return "browser window is inactive while filling Google Forms answers"

                logger.info("Sending answer letter to Google Forms window")
                if not self.dll.AU3_Send(letter, 1):
                    return f"could not send {letter} to Google Forms"
                rand = random.uniform(0.05, 0.15)
                time.sleep(rand)
                
        time.sleep(2)
        logger.info("Submitting Google Forms")
        if not self.dll.AU3_Send("{TAB}{ENTER}", 0):
            return "could not send keys to submit Google Forms"
                
        return None
    
    def create_form(self, name, title, description, questions):
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
        
        logger.info("Browsing Google Forms create")
        err = self.browser.browse("https://forms.google.com/create")
        if err:
            return err
        
        time.sleep(2)
        logger.info("Toggling browser window fullscreen")
        err = self.browser.toggle_fullscreen()
        if err:
            return err
        
        time.sleep(2)
        logger.info("Clicking on top left")
        if not self.dll.AU3_MouseClick("left", 0, 0, 1, 10):
            return "failed to click on top left"

        time.sleep(2)
        logger.info("Moving to Google Forms name")
        for _ in range(2):
            time.sleep(0.5)
            if not self.dll.AU3_Send("{TAB}", 0):
                return "could not send tab key to move to Google Forms name"
        
        time.sleep(1)
        for letter in name:
            logger.info("Checking if browser window is active")
            if not self.dll.AU3_WinActive(self.browser.window_info, ""):
                return "browser window is inactive while filling Google Forms name"

            logger.info("Sending name letter to Google Forms window")
            if not self.dll.AU3_Send(letter, 1):
                return f"could not send {letter} to Google Forms"
            rand = random.uniform(0.05, 0.1)
            time.sleep(rand)
            
        time.sleep(2)
        logger.info("Moving to Google Forms title")
        for _ in range(18):
            time.sleep(0.5)
            if not self.dll.AU3_Send("{TAB}", 0):
                return "could not send tab key to move to Google Forms title"
        
        time.sleep(1)
        logger.info("Filling Google Forms title")
        if not self.dll.AU3_Send("^a{BACKSPACE}", 0):
            return "could not send keys to delete template Google Forms title"
        for letter in title:
            logger.info("Checking if browser window is active")
            if not self.dll.AU3_WinActive(self.browser.window_info, ""):
                return "browser window is inactive while filling Google Forms title"

            logger.info("Sending title letter to Google Forms window")
            if not self.dll.AU3_Send(letter, 1):
                return f"could not send {letter} to Google Forms"
            rand = random.uniform(0.05, 0.1)
            time.sleep(rand)
            
        time.sleep(2)
        logger.info("Moving to Google Forms description")
        for _ in range(2):
            time.sleep(0.5)
            if not self.dll.AU3_Send("{TAB}", 0):
                return "could not send tab key to move to Google Forms description"
        
        time.sleep(1)
        logger.info("Filling Google Forms description")
        for letter in description:
            logger.info("Checking if browser window is active")
            if not self.dll.AU3_WinActive(self.browser.window_info, ""):
                return "browser window is inactive while filling Google Forms description"

            logger.info("Sending description letter to Google Forms window")
            if not self.dll.AU3_Send(letter, 1):
                return f"could not send {letter} to Google Forms"
            rand = random.uniform(0.05, 0.1)
            time.sleep(rand)
            
        time.sleep(2)
        count = len(questions)
        if count > 1:
            logger.info("Adding question boxes")
            first = True
            for _ in range(count-1):
                logger.info("Checking if browser window is active")
                if not self.dll.AU3_WinActive(self.browser.window_info, ""):
                    return "browser window is inactive while adding Google Forms question box"
                
                logger.info("Moving to question box button")
                if first:
                    logger.info("Moving to question box button")
                    for _ in range(8):
                        time.sleep(0.5)
                        if not self.dll.AU3_Send("+{TAB}", 0):
                            return "could not send keys to move to create question box button"
                    
                    first = False
                else:
                    for _ in range(10):
                        time.sleep(0.5)
                        if not self.dll.AU3_Send("+{TAB}", 0):
                            return "could not send keys to move to create question box button"
                        
                logger.info("Pressing question box button")
                if not self.dll.AU3_Send("{ENTER}", 0):
                    return "could not press question box button"
                        
        time.sleep(2)
        for i, (question_type, question_title, is_required) in enumerate(questions):
            logger.info("Sending question title")
            for letter in question_title:
                logger.info("Checking if browser window is active")
                if not self.dll.AU3_WinActive(self.browser.window_info, ""):
                    return "browser window is inactive while filling question title"

                logger.info("Sending question title letter to Google Forms window")
                if not self.dll.AU3_Send(letter, 1):
                    return f"could not send {letter} to Google Forms"
                rand = random.uniform(0.05, 0.15)
                time.sleep(rand)
                
            logger.info("Moving to question type dropdown")
            for _ in range(3):
                time.sleep(0.5)
                if not self.dll.AU3_Send("{TAB}", 0):
                    return "could not send tab key to move to question type dropdown"
            time.sleep(0.5)
            if not self.dll.AU3_Send("{ENTER}", 0):
                return "could not send enter key to appear question type dropdown"
                
            logger.info("Choosing question type")
            if question_type == "Short answer":
                time.sleep(0.5)
                if not self.dll.AU3_Send("{UP}{UP}", 0):
                    return "could not send up key to move question type dropdown selection"
            elif question_type == "Paragraph":
                time.sleep(0.5)
                if not self.dll.AU3_Send("{UP}", 0):
                    return "could not send up key to move question type dropdown selection"
            
            time.sleep(0.5)
            if not self.dll.AU3_Send("{ENTER}", 0):
                return "could not send enter key to choose question type from dropdown"
            
            logger.info("Moving to required toggle")
            for _ in range(3):
                time.sleep(0.5)
                if not self.dll.AU3_Send("{TAB}", 0):
                    return "could not move to required toggle"
            if is_required:
                if not self.dll.AU3_Send("{SPACE}", 0):
                    return "could not toggle required"
                
            if i < count - 1:
                logger.info("Moving to next question box")
                for _ in range(2):
                    time.sleep(0.5)
                    if not self.dll.AU3_Send("{TAB}", 0):
                        return "could not send tab key to move to next question box"
            else:
                logger.info("Moving to publish button")
                shifttabs = 20 + (count - 1) * 2
                for _ in range(shifttabs):
                    time.sleep(0.5)
                    if not self.dll.AU3_Send("+{TAB}", 0):
                        return "could not send keys to move to publish button"
                
                time.sleep(0.5)
                logger.info("Publishing form")
                if not self.dll.AU3_Send("{ENTER}", 0):
                    return "could not send enter key to press publish button"
                time.sleep(0.5)
                if not self.dll.AU3_Send("{TAB}{TAB}", 0):
                    return "could not send tab key to move to publish confirmation button"
                time.sleep(0.5)
                if not self.dll.AU3_Send("{ENTER}", 0):
                    return "could not send enter key to press publish confirmation button"
            
        time.sleep(10)  
        logger.info("Getting form link")
        for _ in range(2):
            time.sleep(0.5)
            if not self.dll.AU3_Send("+{TAB}", 0):
                return "could not send tab key to move to link button"
        time.sleep(0.5)
        if not self.dll.AU3_Send("{SPACE}", 0):
            return "could not send space key to move to link button"
        
        time.sleep(0.5)
        logger.info("Shortening form link")
        if not self.dll.AU3_Send("{TAB}", 0):
            return "could not send tab key to move to shorten link checkbox"
        time.sleep(0.5)
        if not self.dll.AU3_Send("{SPACE}", 0):
            return "could not send space key to check shorten link checkbox"
        
        time.sleep(0.5)
        logger.info("Copying form link to clipboard")
        for _ in range(2):
            time.sleep(0.5)
            if not self.dll.AU3_Send("{TAB}", 0):
                return "could not send tab key to move to copy link button"
        time.sleep(0.5)
        if not self.dll.AU3_Send("{SPACE}", 0):
            return "could not send space key to copy link to clipboard"
        
        time.sleep(0.5)
        logger.info(f"Form created at {pyperclip.paste()}")
        
        time.sleep(2)
        logger.info("Toggling browser window fullscreen")
        err = self.browser.toggle_fullscreen()
        if err:
            return err
        
        return None