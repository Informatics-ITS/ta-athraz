import time
import random
import os
from configs.logger import logger
from activities.apps.native_apps.base import NativeApp

class MicrosoftPaint(NativeApp):
    def __init__(self):
        super().__init__()
        self.window_info = "[CLASS:MSPaintApp]"
        
    def check_existing_window(self):
        logger.info("Checking existing Microsoft Paint window")
        if self.dll.AU3_WinExists(self.window_info, ""):
            logger.info("Activating Microsoft Paint window")
            self.dll.AU3_WinActivate(self.window_info, "")
            logger.info("Waiting Microsoft Paint window to be active")
            if not self.dll.AU3_WinWaitActive(self.window_info, "", 10):
                return "could not activate Microsoft Paint window"
        else:
            return "Microsoft Paint window didn't exist"
        
        return None
        
    def create_window(self):
        logger.info("Running Microsoft Paint")
        if not self.dll.AU3_Run("C:\\Program Files\\WindowsApps\\Microsoft.Paint_11.2502.161.0_x64__8wekyb3d8bbwe\\PaintApp\\mspaint.exe", "", 1):
            return "could not run Microsoft Paint"

        time.sleep(2)
        logger.info("Checking existing Microsoft Paint window")
        if self.dll.AU3_WinExists(self.window_info, ""):
            logger.info("Getting Microsoft Paint window handle")
            handle = self.dll.AU3_WinGetHandle(self.window_info, "")
            if not handle:
                return "could not get window handle"
            else:
                self.window_info = "[HANDLE:%s]" % f"{handle:08X}" 
            logger.info("Activating Microsoft Paint window")
            self.dll.AU3_WinActivate(self.window_info, "")
            logger.info("Waiting Microsoft Paint window to be active")
            if not self.dll.AU3_WinWaitActive(self.window_info, "", 10):
                return "could not activate Microsoft Paint"
        else:
            return "Microsoft Paint window didn't exist"

        time.sleep(2)
        logger.info("Maximizing Microsoft Paint window")
        if not self.dll.AU3_WinSetState(self.window_info, "", 3):
            return "could not maximize Microsoft Paint window"
        
        return None
    
    def open_file(self, path):
        if not path:
            return "path must be provided"
        
        if not os.path.exists(path):
            return f"file path '{path}' does not exist"
        
        if not os.path.isfile(path):
            return f"path '{path}' is not a file"
        
        logger.info("Opening paint file")
        if not self.dll.AU3_Run(f'C:\\Program Files\\WindowsApps\\Microsoft.Paint_11.2502.161.0_x64__8wekyb3d8bbwe\\PaintApp\\mspaint.exe "{path}"', "", 1):
            return "could not open paint file"

        time.sleep(2)
        logger.info("Checking existing Microsoft Paint window")
        if self.dll.AU3_WinExists(self.window_info, ""):
            logger.info("Getting Microsoft Paint window handle")
            handle = self.dll.AU3_WinGetHandle(self.window_info, "")
            if not handle:
                return "could not get window handle"
            else:
                self.window_info = "[HANDLE:%s]" % f"{handle:08X}" 
            logger.info("Activating Microsoft Paint window")
            self.dll.AU3_WinActivate(self.window_info, "")
            logger.info("Waiting Microsoft Paint window to be active")
            if not self.dll.AU3_WinWaitActive(self.window_info, "", 10):
                return "could not activate Microsoft Paint"
        else:
            return "Microsoft Paint window didn't exist"
        
        time.sleep(2)
        logger.info("Maximizing Microsoft Paint window")
        if not self.dll.AU3_WinSetState(self.window_info, "", 3):
            return "could not maximize Microsoft Paint window"

        return None
    
    def draw(self, points=[], speed = 10):
        err = self.check_existing_window()
        if err:
            return err

        if len(points) < 2:
            return "Need at least two points to draw"
        
        left, top, right, bottom = (32, 250, 1875, 1000)

        for i in range(len(points) - 1):
            x0, y0 = points[i]
            x1, y1 = points[i + 1]
            
            if not (left <= x0 <= right and top <= y0 <= bottom):
                return f"Point ({x0}, {y0}) is outside the canvas"
            if not (left <= x1 <= right and top <= y1 <= bottom):
                return f"Point ({x1}, {y1}) is outside the canvas"
            
            logger.info("Checking if Microsoft Paint window is active")
            if not self.dll.AU3_WinActive(self.window_info, ""):
                break
            
            logger.info(f"Dragging mouse from ({x0}, {y0}) to ({x1}, {y1})")
            success = self.dll.AU3_MouseClickDrag("left", x0, y0, x1, y1, speed)
            if not success:
                return f"Failed to drag from ({x0}, {y0}) to ({x1}, {y1})"

        return None 
    
    def draw_random(self, count = 5, speed = 10):
        err = self.check_existing_window()
        if err:
            return err
        
        if count < 2:
            return "Need at least two points to draw"
        
        left, top, right, bottom = (32, 250, 1875, 1000)
        x0 = random.randint(left, right)
        y0 = random.randint(top, bottom)
        
        for i in range(count - 1):
            x1 = random.randint(left, right)
            y1 = random.randint(top, bottom)
            
            if not (left <= x0 <= right and top <= y0 <= bottom):
                return f"Point ({x0}, {y0}) is outside the canvas"
            if not (left <= x1 <= right and top <= y1 <= bottom):
                return f"Point ({x1}, {y1}) is outside the canvas"
            
            logger.info("Checking if Microsoft Paint window is active")
            if not self.dll.AU3_WinActive(self.window_info, ""):
                break
            
            logger.info(f"Dragging mouse from ({x0}, {y0}) to ({x1}, {y1})")
            success = self.dll.AU3_MouseClickDrag("left", x0, y0, x1, y1, speed)
            if not success:
                return f"Failed to drag from ({x0}, {y0}) to ({x1}, {y1})"
            
            x0 = x1
            y0 = y1

        return None
    
    def change_thickness(self, direction, count):
        if direction != "up" and direction != "down":
            return "invalid change thickness direction"
        
        err = self.check_existing_window()
        if err:
            return err
        
        time.sleep(2)
        logger.info("Sending keys to activate change thickness bar")
        if not self.dll.AU3_Send("{ALT}sz", 0):
            return "could not send keys to change thickness bar"
        
        time.sleep(2)
        for _ in range(count):
            logger.info("Checking if Microsoft Paint window is active")
            if not self.dll.AU3_WinActive(self.window_info, ""):
                break
            
            logger.info("Sending keys to change thickness")
            if direction == "up":
                if not self.dll.AU3_Send("{UP}", 0):
                    return "could not send keys to change thickness"
            else:
                if not self.dll.AU3_Send("{DN}", 0):
                    return "could not send keys to change thickness"
            rand = random.uniform(0.05, 0.15)
            time.sleep(rand)
        
        return None