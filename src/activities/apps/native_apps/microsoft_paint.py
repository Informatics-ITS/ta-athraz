import time
import random
import os
from ctypes import windll
from configs.logger import logger
from activities.apps.native_apps.base import NativeApp

class MicrosoftPaint(NativeApp):
    def __init__(self):
        super().__init__()
        self.window_info = "[CLASS:MSPaintApp]"
        self.image_properties_window_info = "[CLASS:XAMLModalWindow]"
        self.image_width = 0
        self.image_height = 0

    def _get_executable_path(self):
        path = os.path.join(os.environ['USERPROFILE'], 'AppData', 'Local', 'Microsoft', 'WindowsApps', 'mspaint.exe')
        if os.path.exists(path):
            return path
        
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
        logger.info("Getting Microsoft Paint executable path")
        executable_path = self._get_executable_path()
        if not executable_path:
            return "could not get Microsoft Paint executable path"
        
        logger.info("Creating new Microsoft Paint window")
        if not self.dll.AU3_Run(executable_path, "", 1):
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
        
        logger.info("Getting Microsoft Paint executable path")
        executable_path = self._get_executable_path()
        if not executable_path:
            return "could not get Microsoft Paint executable path"
        
        logger.info("Opening paint file")
        if not self.dll.AU3_Run(f'{executable_path} "{path}"', "", 1):
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
    
    def _calculate_image_ltrb(self):
        user32 = windll.user32
        gdi32 = windll.gdi32
        user32.SetProcessDPIAware()
        screen_width = user32.GetSystemMetrics(0)
        screen_height = user32.GetSystemMetrics(1)
        hdc = user32.GetDC(0)
        dpi = gdi32.GetDeviceCaps(hdc, 88)
        user32.ReleaseDC(0, hdc)
        scale = int((dpi / 96) * 100)

        left = (screen_width - self.image_width) / 2
        top = (screen_height - (scale * 2.15) - self.image_height) / 2 + (scale * 1.75)
        right = left + self.image_width
        bottom = top + self.image_height
        
        if left < 37:
            left = 37
        if top < (scale * 1.75 + 37):
            top = (scale * 1.75 + 37)
        if right > screen_width - 37:
            right = screen_width - 37
        if bottom > screen_height - (scale * 0.4):
            bottom = screen_height - (scale * 0.4)

        return int(left), int(top), int(right), int(bottom)
    
    def draw(self, points=[], mouse_speed = 10):
        err = self.check_existing_window()
        if err:
            return err

        if len(points) < 2:
            return "Need at least two points to draw"
        
        ltrb = self._calculate_image_ltrb()
        if ltrb == None:
            return "failed to calculate image left, top, right, and bottom"
        left, top, right, bottom = ltrb

        for i in range(len(points) - 1):
            x0, y0 = points[i]
            x1, y1 = points[i + 1]
            
            if not (left <= x0 <= right and top <= y0 <= bottom):
                return f"Point ({x0}, {y0}) is outside the canvas"
            if not (left <= x1 <= right and top <= y1 <= bottom):
                return f"Point ({x1}, {y1}) is outside the canvas"
            
            logger.info("Checking if Microsoft Paint window is active")
            if not self.dll.AU3_WinActive(self.window_info, ""):
                return "Microsoft Paint window is inactive"
            
            logger.info(f"Dragging mouse from ({x0}, {y0}) to ({x1}, {y1})")
            success = self.dll.AU3_MouseClickDrag("left", x0, y0, x1, y1, mouse_speed)
            if not success:
                return f"Failed to drag from ({x0}, {y0}) to ({x1}, {y1})"

        return None
    
    def draw_random(self, count = 5, mouse_speed = 10):
        err = self.check_existing_window()
        if err:
            return err
        
        if count < 2:
            return "Need at least two points to draw"
        
        ltrb = self._calculate_image_ltrb()
        if ltrb == None:
            return "failed to calculate image left, top, right, and bottom"
        left, top, right, bottom = ltrb
        
        x0 = random.randint(left + 1, right - 1)
        y0 = random.randint(top + 1, bottom - 1)
        
        for i in range(count - 1):
            x1 = random.randint(left + 1, right - 1)
            y1 = random.randint(top + 1, bottom - 1)
            
            logger.info("Checking if Microsoft Paint window is active")
            if not self.dll.AU3_WinActive(self.window_info, ""):
                return "Microsoft Paint window is inactive"
            
            logger.info(f"Dragging mouse from ({x0}, {y0}) to ({x1}, {y1})")
            success = self.dll.AU3_MouseClickDrag("left", x0, y0, x1, y1, mouse_speed)
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
                return "Microsoft Paint window is inactive"
            
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
    
    def change_image_size(self, width, height):
        if not width:
            return "image width must be provided"
        if not height:
            return "image height must be provided"
        
        err = self.check_existing_window()
        if err:
            return err
        
        time.sleep(2)
        logger.info("Sending keys to activate image properties modal")
        if not self.dll.AU3_Send("^e", 0):
            return "could not send keys to activate image properties modal"
        
        time.sleep(1)
        logger.info("Checking existing image properties modal")
        if self.dll.AU3_WinExists(self.image_properties_window_info, ""):
            logger.info("Activating image properties modal")
            self.dll.AU3_WinActivate(self.image_properties_window_info, "")
            logger.info("Waiting image properties modal to be active")
            if not self.dll.AU3_WinWaitActive(self.image_properties_window_info, "", 10):
                return "could not activate image properties modal"
        else:
            return "image properties modal didn't exist"
        
        time.sleep(2)
        logger.info("Sending image width to image properties modal")
        if not self.dll.AU3_Send(str(width), 1):
            return "could not send image width to image properties modal"
        
        time.sleep(2)
        logger.info("Sending tab to image properties modal")
        if not self.dll.AU3_Send("{TAB}", 0):
            return "could not send tab to image properties modal"
        
        time.sleep(2)
        logger.info("Sending image height to image properties modal")
        if not self.dll.AU3_Send(str(height), 1):
            return "could not send image height to image properties modal"
        
        time.sleep(1)
        logger.info("Sending enter to image properties modal")
        if not self.dll.AU3_Send("{ENTER}", 0):
            return "could not send Enter key to change image size"
        
        self.image_width = width
        self.image_height = height
        
        return None
    
    def save_file(self):
        err = self.check_existing_window()
        if err:
            return err
        
        logger.info("Saving Paint file")
        if not self.dll.AU3_Send("^s", 0):
            return "could not send ctrl+s to save file"
        
        return None