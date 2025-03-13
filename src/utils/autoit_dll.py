import ctypes

class AutoItDLL():
    def __init__(self):
        self.dll = ctypes.WinDLL(r"C:\Program Files (x86)\AutoIt3\AutoItX\AutoItX3_X64.dll")
        self.dll.AU3_Init()