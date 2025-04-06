from abc import ABC
from utils.autoit_dll import AutoItDLL

class NativeApp(ABC):
    def __init__(self):
        self.dll = AutoItDLL().dll