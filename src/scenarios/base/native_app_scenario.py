from abc import ABC, abstractmethod
from utils.autoit_dll import AutoItDLL

class NativeAppScenario(ABC):
    def __init__(self):
        self.dll = AutoItDLL().dll
        
    @abstractmethod
    def run(self):
        pass