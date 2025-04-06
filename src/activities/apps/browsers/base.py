from abc import ABC, abstractmethod
from utils.autoit_dll import AutoItDLL

class Browser(ABC):
    def __init__(self):
        self.dll = AutoItDLL().dll
        
    @abstractmethod
    def create_window(self):
        pass
    
    @abstractmethod
    def browse(self, url: str):
        pass