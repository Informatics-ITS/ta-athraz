import unittest
from unittest.mock import MagicMock
from activities.others.autoit_function import AutoItFunction

class TestAutoItFunction(unittest.TestCase):
    def setUp(self):
        self.autoit_function = AutoItFunction()
        self.autoit_function.dll = MagicMock()
        
    def test_mouse_click_success(self):
        button = "LEFT"
        x = 100
        y = 200
        clicks = 1
        speed = 5
        self.autoit_function.dll.AU3_MouseClick.return_value = 1

        res = self.autoit_function.mouse_click(button=button, x=x, y=y, clicks=clicks, speed=speed)

        self.autoit_function.dll.AU3_MouseClick.assert_called_once_with(button, x, y, clicks, speed)
        self.assertIsNone(res)
        
    def test_mouse_click_failed(self):
        button = "LEFT"
        x = 100
        y = 200
        clicks = 1
        speed = 5
        self.autoit_function.dll.AU3_MouseClick.return_value = 0

        res = self.autoit_function.mouse_click(button=button, x=x, y=y, clicks=clicks, speed=speed)

        self.autoit_function.dll.AU3_MouseClick.assert_called_once_with(button, x, y, clicks, speed)
        self.assertEqual(res, f"Failed to click mouse at ({x}, {y})")
        
    def test_mouse_click_drag_success(self):
        button = "LEFT"
        x1 = 100
        y1 = 200
        x2 = 300
        y2 = 400
        speed = 10
        self.autoit_function.dll.AU3_MouseClickDrag.return_value = 1
        
        res = self.autoit_function.mouse_click_drag(button=button, x1=x1, y1=y1, x2=x2, y2=y2, speed=speed)
        
        self.autoit_function.dll.AU3_MouseClickDrag.assert_called_once_with(button, x1, y1, x2, y2, speed)
        self.assertIsNone(res)
        
    def test_mouse_click_drag_failed(self):
        button = "LEFT"
        x1 = 100
        y1 = 200
        x2 = 300
        y2 = 400
        speed = 10
        self.autoit_function.dll.AU3_MouseClickDrag.return_value = 0
        
        res = self.autoit_function.mouse_click_drag(button=button, x1=x1, y1=y1, x2=x2, y2=y2, speed=speed)
        
        self.autoit_function.dll.AU3_MouseClickDrag.assert_called_once_with(button, x1, y1, x2, y2, speed)
        self.assertEqual(res, f"Failed to drag mouse from ({x1}, {y1}) to ({x2}, {y2})")
        
    def test_mouse_down_success(self):
        button = "LEFT"
        self.autoit_function.dll.AU3_MouseDown.return_value = 1
        
        res = self.autoit_function.mouse_down(button=button)
        
        self.autoit_function.dll.AU3_MouseDown.assert_called_once_with(button)
        self.assertIsNone(res)
        
    def test_mouse_down_failed(self):
        button = "LEFT"
        self.autoit_function.dll.AU3_MouseDown.return_value = 0
        
        res = self.autoit_function.mouse_down(button=button)
        
        self.autoit_function.dll.AU3_MouseDown.assert_called_once_with(button)
        self.assertEqual(res, "Failed to perform mouse down event")
        
    def test_mouse_move_success(self):
        x = 100
        y = 200
        speed = 10
        self.autoit_function.dll.AU3_MouseMove.return_value = 1
        
        res = self.autoit_function.mouse_move(x=x, y=y, speed=speed)
        
        self.autoit_function.dll.AU3_MouseMove.assert_called_once_with(x, y, speed)
        self.assertIsNone(res)
        
    def test_mouse_move_failed(self):
        x = 100
        y = 200
        speed = 10
        self.autoit_function.dll.AU3_MouseMove.return_value = 0
        
        res = self.autoit_function.mouse_move(x=x, y=y, speed=speed)
        
        self.autoit_function.dll.AU3_MouseMove.assert_called_once_with(x, y, speed)
        self.assertEqual(res, f"Failed to move mouse to ({x}, {y})")
        
    def test_mouse_up_success(self):
        button = "LEFT"
        self.autoit_function.dll.AU3_MouseUp.return_value = 1
        
        res = self.autoit_function.mouse_up(button=button)
        
        self.autoit_function.dll.AU3_MouseUp.assert_called_once_with(button)
        self.assertIsNone(res)
        
    def test_mouse_up_failed(self):
        button = "LEFT"
        self.autoit_function.dll.AU3_MouseUp.return_value = 0
        
        res = self.autoit_function.mouse_up(button=button)
        
        self.autoit_function.dll.AU3_MouseUp.assert_called_once_with(button)
        self.assertEqual(res, "Failed to perform mouse up event")
        
    def test_mouse_wheel_success(self):
        direction = "DOWN"
        clicks = 30
        self.autoit_function.dll.AU3_MouseWheel.return_value = 1
        
        res = self.autoit_function.mouse_wheel(direction=direction, clicks=clicks)
        
        self.autoit_function.dll.AU3_MouseWheel.assert_called_once_with(direction, clicks)
        self.assertIsNone(res)
        
    def test_mouse_wheel_failed(self):
        direction = "DOWN"
        clicks = 30
        self.autoit_function.dll.AU3_MouseWheel.return_value = 0
        
        res = self.autoit_function.mouse_wheel(direction=direction, clicks=clicks)
        
        self.autoit_function.dll.AU3_MouseWheel.assert_called_once_with(direction, clicks)
        self.assertEqual(res, f"Falied to scroll mouse wheel {direction}")
        
    def test_run_success(self):
        exe_path = "path/to/exe"
        working_dir = ""
        flag = 1
        self.autoit_function.dll.AU3_Run.return_value = 1
        
        res = self.autoit_function.run(exe_path=exe_path, working_dir=working_dir, flag=flag)
        
        self.autoit_function.dll.AU3_Run.assert_called_once_with(exe_path, working_dir, flag)
        self.assertIsNone(res)
        
    def test_run_failed(self):
        exe_path = "path/to/exe"
        working_dir = ""
        flag = 1
        self.autoit_function.dll.AU3_Run.return_value = 0
        
        res = self.autoit_function.run(exe_path=exe_path, working_dir=working_dir, flag=flag)
        
        self.autoit_function.dll.AU3_Run.assert_called_once_with(exe_path, working_dir, flag)
        self.assertEqual(res, f"Failed to run: {exe_path}")
        
    def test_send_success(self):
        text = "lorem ipsum"
        mode = 1
        self.autoit_function.dll.AU3_Send.return_value = 1
        
        res = self.autoit_function.send(text=text, mode=mode)
        
        self.autoit_function.dll.AU3_Send.assert_called_once_with(text, mode)
        self.assertIsNone(res)
        
    def test_send_failed(self):
        text = "lorem ipsum"
        mode = 1
        self.autoit_function.dll.AU3_Send.return_value = 0
        
        res = self.autoit_function.send(text=text, mode=mode)
        
        self.autoit_function.dll.AU3_Send.assert_called_once_with(text, mode)
        self.assertEqual(res,f"Failed to send text: {text}")
        
    def test_win_activate_success(self):
        title = "[CLASS:Notepad]"
        text = ""
        self.autoit_function.dll.AU3_WinActivate.return_value = 1
        
        res = self.autoit_function.win_activate(title=title, text=text)
        
        self.autoit_function.dll.AU3_WinActivate.assert_called_once_with(title, text)
        self.assertIsNone(res)
        
    def test_win_activate_failed(self):
        title = "[CLASS:Notepad]"
        text = ""
        self.autoit_function.dll.AU3_WinActivate.return_value = 0
        
        res = self.autoit_function.win_activate(title=title, text=text)
        
        self.autoit_function.dll.AU3_WinActivate.assert_called_once_with(title, text)
        self.assertEqual(res, f"Failed to activate window {title}")
        
    def test_win_active_success(self):
        title = "[CLASS:Notepad]"
        text = ""
        self.autoit_function.dll.AU3_WinActive.return_value = 1
        
        res = self.autoit_function.win_active(title=title, text=text)
        
        self.autoit_function.dll.AU3_WinActive.assert_called_once_with(title, text)
        self.assertIsNone(res)
        
    def test_win_active_failed(self):
        title = "[CLASS:Notepad]"
        text = ""
        self.autoit_function.dll.AU3_WinActive.return_value = 0
        
        res = self.autoit_function.win_active(title=title, text=text)
        
        self.autoit_function.dll.AU3_WinActive.assert_called_once_with(title, text)
        self.assertEqual(res, f"Window {title} is not active")
        
    def test_win_close_success(self):
        title = "[CLASS:Notepad]"
        text = ""
        self.autoit_function.dll.AU3_WinClose.return_value = 1
        
        res = self.autoit_function.win_close(title=title, text=text)
        
        self.autoit_function.dll.AU3_WinClose.assert_called_once_with(title, text)
        self.assertIsNone(res)
        
    def test_win_close_failed(self):
        title = "[CLASS:Notepad]"
        text = ""
        self.autoit_function.dll.AU3_WinClose.return_value = 0
        
        res = self.autoit_function.win_close(title=title, text=text)
        
        self.autoit_function.dll.AU3_WinClose.assert_called_once_with(title, text)
        self.assertEqual(res, f"Failed to close window {title}")

    def test_win_exists_success(self):
        title = "[CLASS:Notepad]"
        text = ""
        self.autoit_function.dll.AU3_WinExists.return_value = 1
        
        res = self.autoit_function.win_exists(title=title, text=text)
        
        self.autoit_function.dll.AU3_WinExists.assert_called_once_with(title, text)
        self.assertIsNone(res)
        
    def test_win_exists_failed(self):
        title = "[CLASS:Notepad]"
        text = ""
        self.autoit_function.dll.AU3_WinExists.return_value = 0
        
        res = self.autoit_function.win_exists(title=title, text=text)
        
        self.autoit_function.dll.AU3_WinExists.assert_called_once_with(title, text)
        self.assertEqual(res, f"Failed to check if window {title} exists")
        
    def test_win_set_state_success(self):
        title = "[CLASS:Notepad]"
        text = ""
        flag = 3
        self.autoit_function.dll.AU3_WinSetState.return_value = 1
        
        res = self.autoit_function.win_set_state(title=title, text=text, flag=flag)
        
        self.autoit_function.dll.AU3_WinSetState.assert_called_once_with(title, text, flag)
        self.assertIsNone(res)
        
    def test_win_set_state_failed(self):
        title = "[CLASS:Notepad]"
        text = ""
        flag = 3
        self.autoit_function.dll.AU3_WinSetState.return_value = 0
        
        res = self.autoit_function.win_set_state(title=title, text=text, flag=flag)
        
        self.autoit_function.dll.AU3_WinSetState.assert_called_once_with(title, text, flag)
        self.assertEqual(res, f"Failed to set window state for {title}")
        
    def test_win_wait_success(self):
        title = "[CLASS:Notepad]"
        text = ""
        timeout = 10        
        self.autoit_function.dll.AU3_WinWait.return_value = 1
        
        res = self.autoit_function.win_wait(title=title, text=text, timeout=timeout)
        
        self.autoit_function.dll.AU3_WinWait.assert_called_once_with(title, text, timeout)
        self.assertIsNone(res)
        
    def test_win_wait_failed(self):
        title = "[CLASS:Notepad]"
        text = ""
        timeout = 10        
        self.autoit_function.dll.AU3_WinWait.return_value = 0
        
        res = self.autoit_function.win_wait(title=title, text=text, timeout=timeout)
        
        self.autoit_function.dll.AU3_WinWait.assert_called_once_with(title, text, timeout)
        self.assertEqual(res, f"Timed out waiting for window {title} to exist")
        
    def test_win_wait_active_success(self):
        title = "[CLASS:Notepad]"
        text = ""
        timeout = 10        
        self.autoit_function.dll.AU3_WinWaitActive.return_value = 1
        
        res = self.autoit_function.win_wait_active(title=title, text=text, timeout=timeout)
        
        self.autoit_function.dll.AU3_WinWaitActive.assert_called_once_with(title, text, timeout)
        self.assertIsNone(res)
        
    def test_win_wait_active_failed(self):
        title = "[CLASS:Notepad]"
        text = ""
        timeout = 10        
        self.autoit_function.dll.AU3_WinWaitActive.return_value = 0
        
        res = self.autoit_function.win_wait_active(title=title, text=text, timeout=timeout)
        
        self.autoit_function.dll.AU3_WinWaitActive.assert_called_once_with(title, text, timeout)
        self.assertEqual(res, f"Timed out waiting for window {title} to active")
        
    def test_win_wait_close_success(self):
        title = "[CLASS:Notepad]"
        text = ""
        timeout = 10        
        self.autoit_function.dll.AU3_WinWaitClose.return_value = 1
        
        res = self.autoit_function.win_wait_close(title=title, text=text, timeout=timeout)
        
        self.autoit_function.dll.AU3_WinWaitClose.assert_called_once_with(title, text, timeout)
        self.assertIsNone(res)
        
    def test_win_wait_close_failed(self):
        title = "[CLASS:Notepad]"
        text = ""
        timeout = 10        
        self.autoit_function.dll.AU3_WinWaitClose.return_value = 0
        
        res = self.autoit_function.win_wait_close(title=title, text=text, timeout=timeout)
        
        self.autoit_function.dll.AU3_WinWaitClose.assert_called_once_with(title, text, timeout)
        self.assertEqual(res, f"Timed out waiting for window {title} to close")