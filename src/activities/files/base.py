from abc import ABC, abstractmethod
from utils.autoit_dll import AutoItDLL

class File(ABC):
    def __init__(self):
        self.dll = AutoItDLL().dll
        
    @abstractmethod
    def _check_existing_window(self):
        pass
        
    @abstractmethod
    def create_window(self):
        pass
        
    @abstractmethod
    def open_file(self, path: str):
        pass
        
    @abstractmethod
    def change_directory(self, path: str):
        pass
    
    @abstractmethod
    def create_directory(self, parent_path: str, dir_name: str):
        pass