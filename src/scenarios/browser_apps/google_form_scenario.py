import time
import random
from configs.logger import logger
from scenarios.base.browser_app_scenario import BrowserAppScenario

class GoogleFormScenario(BrowserAppScenario):        
    def fill(self, url, answers):
        print("url", url)
        if not url.startswith("https://forms.gle/") and not url.startswith("https://https://docs.google.com/forms/"):
            return "Invalid Google Form URL"
        
        self.browser.browse(url)
        
        time.sleep(2)
        first = True
        for answer in answers:
            logger.info("Moving to Google Form question")
            if first:
                if not self.dll.AU3_Send("{TAB}{TAB}{TAB}{TAB}", 0):
                    return "could not move to Google Form question"
                first = False
            else:
                if not self.dll.AU3_Send("{TAB}", 0):
                    return "could not move to next question"
            
            logger.info("Sending answer to Google Form")
            for letter in answer:
                logger.info("Checking if Google Form window is active")
                if not self.dll.AU3_WinActive(self.browser.window_info, ""):
                    break

                logger.info("Sending letter to Google Form window")
                if not self.dll.AU3_Send(letter, 1):
                    return f"could not send {letter} to Google Form"
                rand = random.uniform(0.05, 0.15)
                time.sleep(rand)
                
        time.sleep(2)
        logger.info("Submitting Google Form")
        if not self.dll.AU3_Send("{TAB}{ENTER}", 0):
            return "could not submit Google Form"
                
        return None