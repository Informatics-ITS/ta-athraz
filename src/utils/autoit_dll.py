import ctypes
import os

class AutoItDLL():
    def __init__(self):
        dll_path = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "libs", "AutoItX3_X64.dll"))
        self.dll = ctypes.WinDLL(dll_path)
        self.dll.AU3_Init()