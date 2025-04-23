from abc import ABC, abstractmethod
from utils.autoit_dll import AutoItDLL

class Browser(ABC):
    def __init__(self):
        self.dll = AutoItDLL().dll
        
    @abstractmethod
    def check_existing_window(self):
        pass
        
    @abstractmethod
    def create_window(self):
        pass
    
    @abstractmethod
    def create_tab(self):
        pass
    
    @abstractmethod
    def browse(self, url: str):
        pass
    
    @abstractmethod
    def scroll(self, direction: str, clicks: int, scroll_delay: float):
        pass
    
    @abstractmethod
    def zoom_in(self, count: int):
        pass
    
    @abstractmethod
    def zoom_out(self, count: int):
        pass