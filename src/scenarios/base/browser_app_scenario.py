from abc import ABC, abstractmethod
from scenarios.base.browser_scenario import BrowserScenario
from utils.autoit_dll import AutoItDLL

class BrowserAppScenario(ABC):
    def __init__(self, browser: BrowserScenario):
        self.browser = browser
        self.dll = AutoItDLL().dll