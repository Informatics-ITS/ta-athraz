from abc import ABC
from activities.apps.browsers.base import Browser
from utils.autoit_dll import AutoItDLL

class BrowserApp(ABC):
    def __init__(self, browser: Browser):
        self.browser = browser
        self.dll = AutoItDLL().dll