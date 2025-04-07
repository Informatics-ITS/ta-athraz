from abc import ABC, abstractmethod
from utils.autoit_dll import AutoItDLL

class NativeApp(ABC):
    def __init__(self):
        self.dll = AutoItDLL().dll
        
    @abstractmethod
    def _check_existing_window(self):
        pass