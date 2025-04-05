from abc import ABC, abstractmethod
from utils.autoit_dll import AutoItDLL

class BrowserScenario(ABC):
    def __init__(self):
        self.dll = AutoItDLL().dll
        
    @abstractmethod
    def run(self):
        pass
    
    @abstractmethod
    def browse(self, url: str):
        pass