import time
import random
import pyperclip
from configs.logger import logger
from scenarios.base.browser_app_scenario import BrowserAppScenario

class GoogleFormScenario(BrowserAppScenario):        
    def fill(self, url, answers):
        if not url.startswith("https://forms.gle/") and not url.startswith("https://https://docs.google.com/forms/"):
            return "Invalid Google Form URL"
        
        logger.info("Browsing Google Form URL")
        err = self.browser.browse(url)
        if err:
            return err
        
        time.sleep(2)
        first = True
        for answer in answers:
            logger.info("Checking if browser window is active")
            if not self.dll.AU3_WinActive(self.browser.window_info, ""):
                return "browser window is inactive while moving to Google Form questions"
            
            if first:
                logger.info("Moving to Google Form question")
                if not self.dll.AU3_Send("{TAB}{TAB}{TAB}", 0):
                    return "could not send tab key to move to Google Form question"
                first = False
            else:
                logger.info("Moving to next question")
                if not self.dll.AU3_Send("{TAB}", 0):
                    return "could not send tab key to move to next question"

            time.sleep(random.uniform(1, 3))
            
            logger.info("Sending answer to Google Form")
            for letter in answer:
                logger.info("Checking if browser window is active")
                if not self.dll.AU3_WinActive(self.browser.window_info, ""):
                    return "browser window is inactive while filling Google Form answers"

                logger.info("Sending answer letter to Google Form window")
                if not self.dll.AU3_Send(letter, 1):
                    return f"could not send {letter} to Google Form"
                rand = random.uniform(0.05, 0.15)
                time.sleep(rand)
                
        time.sleep(2)
        logger.info("Submitting Google Form")
        if not self.dll.AU3_Send("{TAB}{ENTER}", 0):
            return "could not send keys to submit Google Form"
                
        return None
    
    def create(self, title, description, questions):        
        logger.info("Browsing Google Form create")
        err = self.browser.browse("https://forms.google.com/create")
        if err:
            return err
        
        time.sleep(6)
        logger.info("Moving to Google Form title")
        for _ in range(10):
            time.sleep(0.5)
            if not self.dll.AU3_Send("{TAB}", 0):
                return "could not send tab key to move to Google Form title"
        
        time.sleep(1)
        logger.info("Filling Google Form title")
        if not self.dll.AU3_Send("^a{BACKSPACE}", 0):
            return "could not send keys to delete template Google Form title"
        for letter in title:
            logger.info("Checking if browser window is active")
            if not self.dll.AU3_WinActive(self.browser.window_info, ""):
                return "browser window is inactive while filling Google Form title"

            logger.info("Sending title letter to Google Form window")
            if not self.dll.AU3_Send(letter, 1):
                return f"could not send {letter} to Google Form"
            rand = random.uniform(0.05, 0.1)
            time.sleep(rand)
            
        time.sleep(2)
        logger.info("Moving to Google Form description")
        for _ in range(2):
            time.sleep(0.5)
            if not self.dll.AU3_Send("{TAB}", 0):
                return "could not send tab key to move to Google Form description"
        
        time.sleep(1)
        logger.info("Filling Google Form description")
        for letter in description:
            logger.info("Checking if browser window is active")
            if not self.dll.AU3_WinActive(self.browser.window_info, ""):
                return "browser window is inactive while filling Google Form description"

            logger.info("Sending description letter to Google Form window")
            if not self.dll.AU3_Send(letter, 1):
                return f"could not send {letter} to Google Form"
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
                    return "browser window is inactive while adding Google Form question box"
                
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

                logger.info("Sending question title letter to Google Form window")
                if not self.dll.AU3_Send(letter, 1):
                    return f"could not send {letter} to Google Form"
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
        
        return None