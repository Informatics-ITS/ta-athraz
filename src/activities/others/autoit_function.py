from configs.logger import logger
from utils.autoit_dll import AutoItDLL

class AutoItFunction():
    def __init__(self):
        self.dll = AutoItDLL().dll
    
    def mouse_click(self, button="LEFT", x=-2147483647, y=-2147483647, clicks=1, speed=-1):
        logger.info(f"Clicking mouse button '{button}' at ({x}, {y}) with {clicks} click(s) and speed {speed}")
        if not self.dll.AU3_MouseClick(button, x, y, clicks, speed):
            return f"Failed to click mouse at ({x}, {y})"
        return None
        
    def mouse_click_drag(self, button, x1, y1, x2, y2, speed=-1):
        logger.info(f"Dragging mouse button '{button}' from ({x1}, {y1}) to ({x2}, {y2}) with speed {speed}")
        if not self.dll.AU3_MouseClickDrag(button, x1, y1, x2, y2, speed):
            return f"Failed to drag mouse from ({x1}, {y1}) to ({x2}, {y2})"
        return None
    
    def mouse_down(self, button="LEFT"):
        logger.info(f"Peforming mouse down event from current position")
        if not self.dll.AU3_MouseDown(button):
            return f"Failed to perform mouse down event"
        return None

    def mouse_move(self, x, y, speed=-1):
        logger.info(f"Moving mouse to ({x}, {y}) with speed {speed}")
        if not self.dll.AU3_MouseMove(x, y, speed):
            return f"Failed to move mouse to ({x}, {y})"
        return None
    
    def mouse_up(self, button="LEFT"):
        logger.info(f"Peforming mouse up event from current position")
        if not self.dll.AU3_MouseUp(button):
            return f"Failed to perform mouse up event"
        return None
        
    def mouse_wheel(self, direction, clicks):
        logger.info(f"Scrolling mouse wheel {direction} with {clicks} click(s)")
        if not self.dll.AU3_MouseWheel(direction, clicks):
            return f"Falied to scroll mouse wheel {direction}"
        return None
    
    def run(self, exe_path, working_dir="", flag=1):
        logger.info(f"Running executable: {exe_path}")
        if not self.dll.AU3_Run(exe_path, working_dir, flag):
            return f"Failed to run: {exe_path}"
        return None
    
    def send(self, text, mode=0):
        logger.info(f"Sending text: {text}")
        if not self.dll.AU3_Send(text, mode):
            return f"Failed to send text: {text}"
        return None
    
    def win_activate(self, title, text):
        logger.info(f"Activating window {title}")
        if not self.dll.AU3_WinActivate(title, text):
            return f"Failed to activate window {title}"
        return None
     
    def win_active(self, title, text=""):
        logger.info(f"Checking if window {title} is active")
        if not self.dll.AU3_WinActive(title, text):
            return f"Window {title} is not active"
        return None
        
    def win_close(self, title, text=""):
        logger.info(f"Closing window {title}")
        if not self.dll.AU3_WinClose(title, text):
            return f"Failed to close window {title}"
        return None
        
    def win_exists(self, title, text=""):
        logger.info(f"Checking if window {title} exists")
        if not self.dll.AU3_WinExists(title, text):
            return f"Failed to check if window {title} exists"
        return None

    def win_set_state(self, title, text, flag):
        logger.info(f"Setting window state for {title} with flags {flag}")
        if not self.dll.AU3_WinSetState(title, text, flag):
            return f"Failed to set window state for {title}"
        return None
    
    def win_wait(self, title, text="", timeout=0):
        logger.info(f"Waiting until window {title} exists (timeout={timeout}s)")
        if not self.dll.AU3_WinWait(title, text, timeout):
            return f"Timed out waiting for window {title} to exist"
        return None
    
    def win_wait_active(self, title, text="", timeout=0):
        logger.info(f"Waiting until window {title} is active (timeout={timeout}s)")
        if not self.dll.AU3_WinWaitActive(title, text, timeout):
            return f"Timed out waiting for window {title} to active"
        return None
        
    def win_wait_close(self, title, text="", timeout=0):
        logger.info(f"Waiting until window {title} is closed (timeout={timeout}s)")
        if not self.dll.AU3_WinWaitClose(title, text, timeout):
            return f"Timed out waiting for window {title} to close"
        return None