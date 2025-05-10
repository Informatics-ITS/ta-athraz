import unittest
from unittest.mock import patch, call, MagicMock
from activities.apps.native_apps.microsoft_paint import MicrosoftPaint
from itertools import cycle

class TestMicrosoftPaint(unittest.TestCase):
    def setUp(self):
        self.paint = MicrosoftPaint()
        self.paint.dll = MagicMock()

    def test_check_existing_window_success(self):
        self.paint.dll.AU3_WinExists.return_value = 1
        self.paint.dll.AU3_WinWaitActive.return_value = 1

        res = self.paint.check_existing_window()

        self.paint.dll.AU3_WinExists.assert_called_once_with(self.paint.window_info, "")
        self.paint.dll.AU3_WinActivate.assert_called_once_with(self.paint.window_info, "")
        self.paint.dll.AU3_WinWaitActive.assert_called_once_with(self.paint.window_info, "", 10)
        
        self.assertIsNone(res)
        
    def test_check_existing_window_failed_window_didnt_exist(self):
        self.paint.dll.AU3_WinExists.return_value = 0
        res = self.paint.check_existing_window()
        
        self.paint.dll.AU3_WinExists.assert_called_once_with(self.paint.window_info, "")
        self.paint.dll.AU3_WinActivate.assert_not_called()
        self.paint.dll.AU3_WinWaitActive.assert_not_called()
        
        self.assertEqual(res, "Microsoft Paint window didn't exist")
        
    def test_check_existing_window_failed_to_wait_window_activation(self):
        self.paint.dll.AU3_WinExists.return_value = 1
        self.paint.dll.AU3_WinActivate.return_value = 1
        self.paint.dll.AU3_WinWaitActive.return_value = 0
        
        res = self.paint.check_existing_window()

        self.paint.dll.AU3_WinExists.assert_called_once_with(self.paint.window_info, "")
        self.paint.dll.AU3_WinActivate.assert_called_once_with(self.paint.window_info, "")
        self.paint.dll.AU3_WinWaitActive.assert_called_once_with(self.paint.window_info, "", 10)
        
        self.assertEqual(res, "could not activate Microsoft Paint window")

    @patch("time.sleep", return_value=None)
    @patch.object(MicrosoftPaint, "_get_executable_path", return_value="/path/to/paint/executable")
    def test_create_window_success(self, mock_get_executable_path, mock_sleep):
        self.paint.dll.AU3_Run.return_value = 1
        self.paint.dll.AU3_WinExists.return_value = 1
        self.paint.dll.AU3_WinGetHandle.return_value = 1234567890
        self.paint.dll.AU3_WinActivate.return_value = 1
        self.paint.dll.AU3_WinWaitActive.return_value = 1
        self.paint.dll.AU3_WinSetState.return_value = 1

        res = self.paint.create_window()
        
        mock_get_executable_path.assert_called_once()
        self.paint.dll.AU3_Run.assert_called_once_with("/path/to/paint/executable", "", 1)
        self.paint.dll.AU3_WinExists.assert_called_once()
        self.paint.dll.AU3_WinGetHandle.assert_called_once()
        self.paint.dll.AU3_WinActivate.assert_called_once_with(self.paint.window_info, "")
        self.paint.dll.AU3_WinWaitActive.assert_called_once_with(self.paint.window_info, "", 10)
        self.paint.dll.AU3_WinSetState.assert_called_once_with(self.paint.window_info, "", 3)

        self.assertEqual(mock_sleep.call_count, 2)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    @patch.object(MicrosoftPaint, "_get_executable_path", return_value=None)
    def test_create_window_failed_to_get_executable_path(self, mock_get_executable_path, mock_sleep):
        res = self.paint.create_window()
        
        mock_get_executable_path.assert_called_once()
        self.paint.dll.AU3_Run.assert_not_called()
        self.paint.dll.AU3_WinExists.assert_not_called()
        self.paint.dll.AU3_WinGetHandle.assert_not_called()
        self.paint.dll.AU3_WinActivate.assert_not_called()
        self.paint.dll.AU3_WinWaitActive.assert_not_called()
        self.paint.dll.AU3_WinSetState.assert_not_called()

        self.assertEqual(res, "could not get Microsoft Paint executable path")
        
    @patch("time.sleep", return_value=None)
    @patch.object(MicrosoftPaint, "_get_executable_path", return_value="/path/to/paint/executable")
    def test_create_window_failed_to_run(self, mock_get_executable_path, mock_sleep):
        self.paint.dll.AU3_Run.return_value = 0
        res = self.paint.create_window()
        
        mock_get_executable_path.assert_called_once()
        self.paint.dll.AU3_Run.assert_called_once_with("/path/to/paint/executable", "", 1)
        self.paint.dll.AU3_WinExists.assert_not_called()
        self.paint.dll.AU3_WinGetHandle.assert_not_called()
        self.paint.dll.AU3_WinActivate.assert_not_called()
        self.paint.dll.AU3_WinWaitActive.assert_not_called()
        self.paint.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(res, "could not run Microsoft Paint")
        
    @patch("time.sleep", return_value=None)
    @patch.object(MicrosoftPaint, "_get_executable_path", return_value="/path/to/paint/executable")
    def test_create_window_failed_window_didnt_exist(self, mock_get_executable_path, mock_sleep):
        self.paint.dll.AU3_Run.return_value = 1
        self.paint.dll.AU3_WinExists.return_value = 0

        res = self.paint.create_window()
        
        mock_get_executable_path.assert_called_once()
        self.paint.dll.AU3_Run.assert_called_once_with("/path/to/paint/executable", "", 1)
        self.paint.dll.AU3_WinExists.assert_called_once()
        self.paint.dll.AU3_WinGetHandle.assert_not_called()
        self.paint.dll.AU3_WinActivate.assert_not_called()
        self.paint.dll.AU3_WinWaitActive.assert_not_called()
        self.paint.dll.AU3_WinSetState.assert_not_called()

        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "Microsoft Paint window didn't exist")
        
    @patch("time.sleep", return_value=None)
    @patch.object(MicrosoftPaint, "_get_executable_path", return_value="/path/to/paint/executable")
    def test_create_window_failed_to_get_handle(self, mock_get_executable_path, mock_sleep):
        self.paint.dll.AU3_Run.return_value = 1
        self.paint.dll.AU3_WinExists.return_value = 1
        self.paint.dll.AU3_WinGetHandle.return_value = 0

        res = self.paint.create_window()
        
        mock_get_executable_path.assert_called_once()
        self.paint.dll.AU3_Run.assert_called_once_with("/path/to/paint/executable", "", 1)
        self.paint.dll.AU3_WinExists.assert_called_once()
        self.paint.dll.AU3_WinGetHandle.assert_called_once()
        self.paint.dll.AU3_WinActivate.assert_not_called()
        self.paint.dll.AU3_WinWaitActive.assert_not_called()
        self.paint.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not get window handle")
        
    @patch("time.sleep", return_value=None)
    @patch.object(MicrosoftPaint, "_get_executable_path", return_value="/path/to/paint/executable")
    def test_create_window_failed_to_wait_window_activation(self, mock_get_executable_path, mock_sleep):
        self.paint.dll.AU3_Run.return_value = 1
        self.paint.dll.AU3_WinExists.return_value = 1
        self.paint.dll.AU3_WinGetHandle.return_value = 1234567890
        self.paint.dll.AU3_WinActivate.return_value = 1
        self.paint.dll.AU3_WinWaitActive.return_value = 0

        res = self.paint.create_window()
        
        mock_get_executable_path.assert_called_once()
        self.paint.dll.AU3_Run.assert_called_once_with("/path/to/paint/executable", "", 1)
        self.paint.dll.AU3_WinExists.assert_called_once()
        self.paint.dll.AU3_WinGetHandle.assert_called_once()
        self.paint.dll.AU3_WinActivate.assert_called_once_with(self.paint.window_info, "")
        self.paint.dll.AU3_WinWaitActive.assert_called_once_with(self.paint.window_info, "", 10)
        self.paint.dll.AU3_WinSetState.assert_not_called()

        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not activate Microsoft Paint")
        
    @patch("time.sleep", return_value=None)
    @patch.object(MicrosoftPaint, "_get_executable_path", return_value="/path/to/paint/executable")
    def test_create_window_failed_to_maximize_window(self, mock_get_executable_path, mock_sleep):
        self.paint.dll.AU3_Run.return_value = 1
        self.paint.dll.AU3_WinExists.return_value = 1
        self.paint.dll.AU3_WinGetHandle.return_value = 1234567890
        self.paint.dll.AU3_WinActivate.return_value = 1
        self.paint.dll.AU3_WinWaitActive.return_value = 1
        self.paint.dll.AU3_WinSetState.return_value = 0

        res = self.paint.create_window()
        
        mock_get_executable_path.assert_called_once()
        self.paint.dll.AU3_Run.assert_called_once_with("/path/to/paint/executable", "", 1)
        self.paint.dll.AU3_WinExists.assert_called_once()
        self.paint.dll.AU3_WinGetHandle.assert_called_once()
        self.paint.dll.AU3_WinActivate.assert_called_once_with(self.paint.window_info, "")
        self.paint.dll.AU3_WinWaitActive.assert_called_once_with(self.paint.window_info, "", 10)
        self.paint.dll.AU3_WinSetState.assert_called_once_with(self.paint.window_info, "", 3)

        self.assertEqual(mock_sleep.call_count, 2)
        self.assertEqual(res, "could not maximize Microsoft Paint window")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(MicrosoftPaint, "_get_executable_path", return_value="/path/to/paint/executable")
    def test_open_file_success(self, mock_get_executable_path, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        self.paint.dll.AU3_Run.return_value = 1
        self.paint.dll.AU3_WinExists.return_value = 1
        self.paint.dll.AU3_WinGetHandle.return_value = 1234567890
        self.paint.dll.AU3_WinActivate.return_value = 1
        self.paint.dll.AU3_WinWaitActive.return_value = 1
        self.paint.dll.AU3_WinSetState.return_value = 1
        
        res = self.paint.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_get_executable_path.assert_called_once()
        self.paint.dll.AU3_Run.assert_called_once_with(f'/path/to/paint/executable "{path}"', "", 1)
        self.paint.dll.AU3_WinExists.assert_called_once()
        self.paint.dll.AU3_WinGetHandle.assert_called_once()
        self.paint.dll.AU3_WinActivate.assert_called_once_with(self.paint.window_info, "")
        self.paint.dll.AU3_WinWaitActive.assert_called_once_with(self.paint.window_info, "", 10)
        self.paint.dll.AU3_WinSetState.assert_called_once_with(self.paint.window_info, "", 3)
        
        self.assertEqual(mock_sleep.call_count, 2)
        self.assertIsNone(res)

    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(MicrosoftPaint, "_get_executable_path")
    def test_open_file_failed_path_is_not_provided(self, mock_get_executable_path, mock_os_path, mock_sleep):
        res = self.paint.open_file(path="")
        
        mock_os_path.exists.assert_not_called()
        mock_os_path.isfile.assert_not_called()
        mock_get_executable_path.assert_not_called()
        self.paint.dll.AU3_Run.assert_not_called()
        self.paint.dll.AU3_WinExists.assert_not_called()
        self.paint.dll.AU3_WinGetHandle.assert_not_called()
        self.paint.dll.AU3_WinActivate.assert_not_called()
        self.paint.dll.AU3_WinWaitActive.assert_not_called()
        self.paint.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 0)
        self.assertEqual(res, "path must be provided")

    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(MicrosoftPaint, "_get_executable_path")  
    def test_open_file_failed_path_didnt_exist(self, mock_get_executable_path, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = False
        
        res = self.paint.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_not_called()
        mock_get_executable_path.assert_not_called()
        self.paint.dll.AU3_Run.assert_not_called()
        self.paint.dll.AU3_WinExists.assert_not_called()
        self.paint.dll.AU3_WinGetHandle.assert_not_called()
        self.paint.dll.AU3_WinActivate.assert_not_called()
        self.paint.dll.AU3_WinWaitActive.assert_not_called()
        self.paint.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 0)
        self.assertEqual(res, f"file path '{path}' does not exist")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(MicrosoftPaint, "_get_executable_path")
    def test_open_file_failed_path_is_not_a_file(self, mock_get_executable_path, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = False
        
        res = self.paint.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_get_executable_path.assert_not_called()
        self.paint.dll.AU3_Run.assert_not_called()
        self.paint.dll.AU3_WinExists.assert_not_called()
        self.paint.dll.AU3_WinGetHandle.assert_not_called()
        self.paint.dll.AU3_WinActivate.assert_not_called()
        self.paint.dll.AU3_WinWaitActive.assert_not_called()
        self.paint.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 0)
        self.assertEqual(res, f"path '{path}' is not a file")
               
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(MicrosoftPaint, "_get_executable_path", return_value=None)
    def test_open_file_failed_to_get_executable_path(self, mock_get_executable_path, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        
        res = self.paint.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_get_executable_path.asset_called_once()
        self.paint.dll.AU3_Run.assert_not_called()
        self.paint.dll.AU3_WinExists.assert_not_called()
        self.paint.dll.AU3_WinGetHandle.assert_not_called()
        self.paint.dll.AU3_WinActivate.assert_not_called()
        self.paint.dll.AU3_WinWaitActive.assert_not_called()
        self.paint.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 0)
        self.assertEqual(res, "could not get Microsoft Paint executable path")
          
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(MicrosoftPaint, "_get_executable_path", return_value="/path/to/paint/executable")
    def test_open_file_failed_to_run(self, mock_get_executable_path, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        self.paint.dll.AU3_Run.return_value = 0
        
        res = self.paint.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_get_executable_path.assert_called_once()
        self.paint.dll.AU3_Run.assert_called_once_with(f'/path/to/paint/executable "{path}"', "", 1)
        self.paint.dll.AU3_WinExists.assert_not_called()
        self.paint.dll.AU3_WinGetHandle.assert_not_called()
        self.paint.dll.AU3_WinActivate.assert_not_called()
        self.paint.dll.AU3_WinWaitActive.assert_not_called()
        self.paint.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 0)
        self.assertEqual(res, "could not open paint file")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(MicrosoftPaint, "_get_executable_path", return_value="/path/to/paint/executable")
    def test_open_file_failed_window_didnt_exist(self, mock_get_executable_path, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        self.paint.dll.AU3_Run.return_value = 1
        self.paint.dll.AU3_WinExists.return_value = 0
        
        res = self.paint.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_get_executable_path.assert_called_once()
        self.paint.dll.AU3_Run.assert_called_once_with(f'/path/to/paint/executable "{path}"', "", 1)
        self.paint.dll.AU3_WinExists.assert_called_once()
        self.paint.dll.AU3_WinGetHandle.assert_not_called()
        self.paint.dll.AU3_WinActivate.assert_not_called()
        self.paint.dll.AU3_WinWaitActive.assert_not_called()
        self.paint.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "Microsoft Paint window didn't exist")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(MicrosoftPaint, "_get_executable_path", return_value="/path/to/paint/executable")
    def test_open_file_failed_to_get_handle(self, mock_get_executable_path, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        self.paint.dll.AU3_Run.return_value = 1
        self.paint.dll.AU3_WinExists.return_value = 1
        self.paint.dll.AU3_WinGetHandle.return_value = 0
        
        res = self.paint.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_get_executable_path.assert_called_once()
        self.paint.dll.AU3_Run.assert_called_once_with(f'/path/to/paint/executable "{path}"', "", 1)
        self.paint.dll.AU3_WinExists.assert_called_once()
        self.paint.dll.AU3_WinGetHandle.assert_called_once()
        self.paint.dll.AU3_WinActivate.assert_not_called()
        self.paint.dll.AU3_WinWaitActive.assert_not_called()
        self.paint.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not get window handle")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(MicrosoftPaint, "_get_executable_path", return_value="/path/to/paint/executable")
    def test_open_file_failed_to_wait_window_activation(self, mock_get_executable_path, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        self.paint.dll.AU3_Run.return_value = 1
        self.paint.dll.AU3_WinExists.return_value = 1
        self.paint.dll.AU3_WinGetHandle.return_value = 1234567890
        self.paint.dll.AU3_WinActivate.return_value = 1
        self.paint.dll.AU3_WinWaitActive.return_value = 0
        
        res = self.paint.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_get_executable_path.assert_called_once()
        self.paint.dll.AU3_Run.assert_called_once_with(f'/path/to/paint/executable "{path}"', "", 1)
        self.paint.dll.AU3_WinExists.assert_called_once()
        self.paint.dll.AU3_WinGetHandle.assert_called_once()
        self.paint.dll.AU3_WinActivate.assert_called_once_with(self.paint.window_info, "")
        self.paint.dll.AU3_WinWaitActive.assert_called_once_with(self.paint.window_info, "", 10)
        self.paint.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not activate Microsoft Paint")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(MicrosoftPaint, "_get_executable_path", return_value="/path/to/paint/executable")
    def test_open_file_failed_to_maximize_window(self, mock_get_executable_path, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        self.paint.dll.AU3_Run.return_value = 1
        self.paint.dll.AU3_WinExists.return_value = 1
        self.paint.dll.AU3_WinGetHandle.return_value = 1234567890
        self.paint.dll.AU3_WinActivate.return_value = 1
        self.paint.dll.AU3_WinWaitActive.return_value = 1
        self.paint.dll.AU3_WinSetState.return_value = 0
        
        res = self.paint.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_get_executable_path.assert_called_once()
        self.paint.dll.AU3_Run.assert_called_once_with(f'/path/to/paint/executable "{path}"', "", 1)
        self.paint.dll.AU3_WinExists.assert_called_once()
        self.paint.dll.AU3_WinGetHandle.assert_called_once()
        self.paint.dll.AU3_WinActivate.assert_called_once_with(self.paint.window_info, "")
        self.paint.dll.AU3_WinWaitActive.assert_called_once_with(self.paint.window_info, "", 10)
        self.paint.dll.AU3_WinSetState.assert_called_once_with(self.paint.window_info, "", 3)
        
        self.assertEqual(mock_sleep.call_count, 2)
        self.assertEqual(res, "could not maximize Microsoft Paint window")
        
    @patch.object(MicrosoftPaint, "_calculate_image_ltrb", return_value=(100, 300, 1900, 1000))
    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None)
    def test_draw_success(self, mock_check_existing_window, mock_calculate_image_ltrb):
        points = [[1110, 625], [1102, 675], [1079, 721], [1043, 760], [997, 787], [945, 798], [892, 793], [844, 770], [808, 735], [787, 690]]
        mouse_speed = 10
        self.paint.dll.AU3_WinActive.return_value = 1
        self.paint.dll.AU3_MouseClickDrag.return_value = 1
        
        res = self.paint.draw(points=points, mouse_speed=mouse_speed)
        
        mock_check_existing_window.assert_called_once()
        mock_calculate_image_ltrb.assert_called_once()
        self.paint.dll.AU3_WinActive.assert_has_calls([call(self.paint.window_info, "")] * (len(points) - 1))
        expected_mouse_click_drag_calls = ()
        for i in range(len(points) - 1):
            x0, y0 = points[i]
            x1, y1 = points[i + 1]
            expected_mouse_click_drag_calls += (
                (call("left", x0, y0, x1, y1, mouse_speed),)
            )
        self.paint.dll.AU3_MouseClickDrag.assert_has_calls(expected_mouse_click_drag_calls)
        
        self.assertEqual(self.paint.dll.AU3_WinActive.call_count, len(points) - 1)
        self.assertEqual(self.paint.dll.AU3_MouseClickDrag.call_count, len(points) - 1)
        self.assertIsNone(res)
        
    @patch.object(MicrosoftPaint, "_calculate_image_ltrb", return_value=(100, 300, 1900, 1000))
    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None)
    def test_draw_failed_not_enough_points(self, mock_check_existing_window, mock_calculate_image_ltrb):
        points = [[1110, 625]]
        mouse_speed = 10
        
        res = self.paint.draw(points=points, mouse_speed=mouse_speed)

        mock_check_existing_window.assert_called_once()
        mock_calculate_image_ltrb.assert_not_called()
        self.paint.dll.AU3_WinActive.assert_not_called()
        self.paint.dll.AU3_MouseClickDrag.assert_not_called()
        self.assertEqual(res, "Need at least two points to draw")
        
    @patch.object(MicrosoftPaint, "_calculate_image_ltrb", return_value=None)
    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None)
    def test_draw_failed_to_calculate_image_ltrb(self, mock_check_existing_window, mock_calculate_image_ltrb):
        points = [[1110, 625], [1102, 675], [1079, 721], [1043, 760], [997, 787], [945, 798], [892, 793], [844, 770], [808, 735], [787, 690]]
        mouse_speed = 10
        
        res = self.paint.draw(points=points, mouse_speed=mouse_speed)
        
        mock_check_existing_window.assert_called_once()
        mock_calculate_image_ltrb.assert_called_once()
        self.paint.dll.AU3_WinActive.assert_not_called()
        self.paint.dll.AU3_MouseClickDrag.assert_not_called()
        self.assertEqual(res, "failed to calculate image left, top, right, and bottom")
        
    @patch.object(MicrosoftPaint, "_calculate_image_ltrb", return_value=(100, 300, 1900, 1000))
    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None)
    def test_draw_failed_point_x0_y0_is_outside_the_canvas(self, mock_check_existing_window, mock_calculate_image_ltrb):
        points = [[2000, 1100], [1102, 675], [1079, 721], [1043, 760], [997, 787], [945, 798], [892, 793], [844, 770], [808, 735], [787, 690]]
        mouse_speed = 10
        
        res = self.paint.draw(points=points, mouse_speed=mouse_speed)
        
        mock_check_existing_window.assert_called_once()
        mock_calculate_image_ltrb.assert_called_once()
        self.paint.dll.AU3_WinActive.assert_not_called()
        self.paint.dll.AU3_MouseClickDrag.assert_not_called()
        self.assertEqual(res, f"Point ({points[0][0]}, {points[0][1]}) is outside the canvas")
        
    @patch.object(MicrosoftPaint, "_calculate_image_ltrb", return_value=(100, 300, 1900, 1000))
    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None)
    def test_draw_failed_point_x1_y1_is_outside_the_canvas(self, mock_check_existing_window, mock_calculate_image_ltrb):
        points = [[1110, 625], [2000, 1100], [1079, 721], [1043, 760], [997, 787], [945, 798], [892, 793], [844, 770], [808, 735], [787, 690]]
        mouse_speed = 10
        
        res = self.paint.draw(points=points, mouse_speed=mouse_speed)
        
        mock_check_existing_window.assert_called_once()
        mock_calculate_image_ltrb.assert_called_once()
        self.paint.dll.AU3_WinActive.assert_not_called()
        self.paint.dll.AU3_MouseClickDrag.assert_not_called()
        self.assertEqual(res, f"Point ({points[1][0]}, {points[1][1]}) is outside the canvas")
        
    @patch.object(MicrosoftPaint, "_calculate_image_ltrb", return_value=(100, 300, 1900, 1000))
    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None)
    def test_draw_failed_window_inactive_while_drawing(self, mock_check_existing_window, mock_calculate_image_ltrb):
        points = [[1110, 625], [1102, 675], [1079, 721], [1043, 760], [997, 787], [945, 798], [892, 793], [844, 770], [808, 735], [787, 690]]
        mouse_speed = 10
        self.paint.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        self.paint.dll.AU3_MouseClickDrag.return_value = 1
        
        res = self.paint.draw(points=points, mouse_speed=mouse_speed)
        
        mock_check_existing_window.assert_called_once()
        mock_calculate_image_ltrb.assert_called_once()
        self.paint.dll.AU3_WinActive.assert_has_calls([call(self.paint.window_info, "")] * 6)
        expected_mouse_click_drag_calls = ()
        for i in range(5):
            x0, y0 = points[i]
            x1, y1 = points[i + 1]
            expected_mouse_click_drag_calls += (
                (call("left", x0, y0, x1, y1, mouse_speed),)
            )
        self.paint.dll.AU3_MouseClickDrag.assert_has_calls(expected_mouse_click_drag_calls)
        
        self.assertEqual(self.paint.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.paint.dll.AU3_MouseClickDrag.call_count, 5)
        self.assertEqual(res, "Microsoft Paint window is inactive")
        
    @patch.object(MicrosoftPaint, "_calculate_image_ltrb", return_value=(100, 300, 1900, 1000))
    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None)
    def test_draw_failed_to_mouse_click_drag_while_drawing(self, mock_check_existing_window, mock_calculate_image_ltrb):
        points = [[1110, 625], [1102, 675], [1079, 721], [1043, 760], [997, 787], [945, 798], [892, 793], [844, 770], [808, 735], [787, 690]]
        mouse_speed = 10
        self.paint.dll.AU3_WinActive.return_value = 1
        self.paint.dll.AU3_MouseClickDrag.side_effect = [1] * 5 + [0]
        
        res = self.paint.draw(points=points, mouse_speed=mouse_speed)
        
        mock_check_existing_window.assert_called_once()
        mock_calculate_image_ltrb.assert_called_once()
        self.paint.dll.AU3_WinActive.assert_has_calls([call(self.paint.window_info, "")] * 6)
        expected_mouse_click_drag_calls = ()
        for i in range(6):
            x0, y0 = points[i]
            x1, y1 = points[i + 1]
            expected_mouse_click_drag_calls += (
                (call("left", x0, y0, x1, y1, mouse_speed),)
            )
        self.paint.dll.AU3_MouseClickDrag.assert_has_calls(expected_mouse_click_drag_calls)
        
        self.assertEqual(self.paint.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.paint.dll.AU3_MouseClickDrag.call_count, 6)
        self.assertEqual(res, f"Failed to drag from ({points[5][0]}, {points[5][1]}) to ({points[6][0]}, {points[6][1]})")
        
    @patch.object(MicrosoftPaint, "_calculate_image_ltrb", return_value=(100, 300, 1900, 1000))
    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None)
    @patch("random.randint")
    def test_draw_random_success(self, mock_random_randint, mock_check_existing_window, mock_calculate_image_ltrb):
        count = 10
        mouse_speed = 10
        points = [1110, 625, 1102, 675, 1079, 721, 1043, 760, 997, 787, 945, 798, 892, 793, 844, 770, 808, 735, 787, 690]
        mock_random_randint.side_effect = cycle(points)
        self.paint.dll.AU3_WinActive.return_value = 1
        self.paint.dll.AU3_MouseClickDrag.return_value = 1
        
        res = self.paint.draw_random(count=count, mouse_speed=mouse_speed)
        
        mock_check_existing_window.assert_called_once()
        mock_calculate_image_ltrb.assert_called_once()
        self.paint.dll.AU3_WinActive.assert_has_calls([call(self.paint.window_info, "")] * (count - 1))
        expected_mouse_click_drag_calls = ()
        for i in range(0, count, 2):
            x0, y0 = points[i], points[i+1]
            x1, y1 = points[i+2], points[i+3]
            expected_mouse_click_drag_calls += (
                (call("left", x0, y0, x1, y1, mouse_speed),)
            )
        self.paint.dll.AU3_MouseClickDrag.assert_has_calls(expected_mouse_click_drag_calls)
        
        self.assertEqual(self.paint.dll.AU3_WinActive.call_count, count - 1)
        self.assertEqual(self.paint.dll.AU3_MouseClickDrag.call_count, count - 1)
        self.assertIsNone(res)

    @patch.object(MicrosoftPaint, "_calculate_image_ltrb", return_value=(100, 300, 1900, 1000))
    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None)
    @patch("random.randint")
    def test_draw_random_failed_not_enough_points(self, mock_random_randint, mock_check_existing_window, mock_calculate_image_ltrb):
        count = 1
        mouse_speed = 10
        
        res = self.paint.draw_random(count=count, mouse_speed=mouse_speed)

        mock_check_existing_window.assert_called_once()
        mock_calculate_image_ltrb.assert_not_called()
        mock_random_randint.assert_not_called()
        self.paint.dll.AU3_WinActive.assert_not_called()
        self.paint.dll.AU3_MouseClickDrag.assert_not_called()
        self.assertEqual(res, "Need at least two points to draw")
        
    @patch.object(MicrosoftPaint, "_calculate_image_ltrb", return_value=None)
    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None)
    @patch("random.randint")
    def test_draw_random_failed_to_calculate_image_ltrb(self, mock_random_randint, mock_check_existing_window, mock_calculate_image_ltrb):
        count = 10
        mouse_speed = 10
        
        res = self.paint.draw_random(count=count, mouse_speed=mouse_speed)
        
        mock_check_existing_window.assert_called_once()
        mock_calculate_image_ltrb.assert_called_once()
        mock_random_randint.assert_not_called()
        self.paint.dll.AU3_WinActive.assert_not_called()
        self.paint.dll.AU3_MouseClickDrag.assert_not_called()
        self.assertEqual(res, "failed to calculate image left, top, right, and bottom")
        
    @patch.object(MicrosoftPaint, "_calculate_image_ltrb", return_value=(100, 300, 1900, 1000))
    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None)
    @patch("random.randint")
    def test_draw_random_failed_window_inactive_while_drawing(self, mock_random_randint, mock_check_existing_window, mock_calculate_image_ltrb):
        count = 10
        mouse_speed = 10
        points = [1110, 625, 1102, 675, 1079, 721, 1043, 760, 997, 787, 945, 798, 892, 793, 844, 770, 808, 735, 787, 690]
        mock_random_randint.side_effect = cycle(points)
        self.paint.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        self.paint.dll.AU3_MouseClickDrag.return_value = 1
        
        res = self.paint.draw_random(count=count, mouse_speed=mouse_speed)
        
        mock_check_existing_window.assert_called_once()
        mock_calculate_image_ltrb.assert_called_once()
        self.paint.dll.AU3_WinActive.assert_has_calls([call(self.paint.window_info, "")] * 6)
        expected_mouse_click_drag_calls = ()
        for i in range(0, 5, 2):
            x0, y0 = points[i], points[i+1]
            x1, y1 = points[i+2], points[i+3]
            expected_mouse_click_drag_calls += (
                (call("left", x0, y0, x1, y1, mouse_speed),)
            )
        self.paint.dll.AU3_MouseClickDrag.assert_has_calls(expected_mouse_click_drag_calls)
        
        self.assertEqual(self.paint.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.paint.dll.AU3_MouseClickDrag.call_count, 5)
        self.assertEqual(res, "Microsoft Paint window is inactive")
        
    @patch.object(MicrosoftPaint, "_calculate_image_ltrb", return_value=(100, 300, 1900, 1000))
    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None)
    @patch("random.randint")
    def test_draw_random_failed_to_mouse_click_drag_while_drawing(self, mock_random_randint, mock_check_existing_window, mock_calculate_image_ltrb):
        count = 10
        mouse_speed = 10
        points = [1110, 625, 1102, 675, 1079, 721, 1043, 760, 997, 787, 945, 798, 892, 793, 844, 770, 808, 735, 787, 690]
        mock_random_randint.side_effect = cycle(points)
        self.paint.dll.AU3_WinActive.return_value = 1
        self.paint.dll.AU3_MouseClickDrag.side_effect = [1] * 5 + [0]
        
        res = self.paint.draw_random(count=count, mouse_speed=mouse_speed)
        
        mock_check_existing_window.assert_called_once()
        mock_calculate_image_ltrb.assert_called_once()
        self.paint.dll.AU3_WinActive.assert_has_calls([call(self.paint.window_info, "")] * 6)
        expected_mouse_click_drag_calls = ()
        for i in range(0, 6, 2):
            x0, y0 = points[i], points[i+1]
            x1, y1 = points[i+2], points[i+3]
            expected_mouse_click_drag_calls += (
                (call("left", x0, y0, x1, y1, mouse_speed),)
            )
        self.paint.dll.AU3_MouseClickDrag.assert_has_calls(expected_mouse_click_drag_calls)
        
        self.assertEqual(self.paint.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.paint.dll.AU3_MouseClickDrag.call_count, 6)
        self.assertEqual(res, f"Failed to drag from ({points[10]}, {points[11]}) to ({points[12]}, {points[13]})")
       
    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_change_thickness_success(self, mock_sleep, mock_check_existing_window):
        direction = "up"
        count = 10
        self.paint.dll.AU3_Send.return_value = 1
        self.paint.dll.AU3_WinActive.return_value = 1
        
        res = self.paint.change_thickness(direction=direction, count=count)
        
        mock_check_existing_window.assert_called_once()
        self.paint.dll.AU3_Send.assert_has_calls(
            [call("{ALT}sz", 0)] + 
            [call("{UP}", 0)] * count
        )
        self.paint.dll.AU3_WinActive.assert_has_calls([call(self.paint.window_info, "")] * count)
        
        self.assertEqual(self.paint.dll.AU3_WinActive.call_count, count)
        self.assertEqual(self.paint.dll.AU3_Send.call_count, count + 1)
        self.assertEqual(mock_sleep.call_count, count + 2)
        self.assertIsNone(res)
        
    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_change_failed_invalid_direction(self, mock_sleep, mock_check_existing_window):
        direction = "abc"
        count = 10
        
        res = self.paint.change_thickness(direction=direction, count=count)
        
        mock_check_existing_window.assert_not_called()
        self.paint.dll.AU3_Send.assert_not_called()
        self.paint.dll.AU3_WinActive.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, "invalid change thickness direction")
        
    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_change_failed_to_send_keys_to_activate_change_thickness_bar(self, mock_sleep, mock_check_existing_window):
        direction = "up"
        count = 10
        self.paint.dll.AU3_Send.return_value = 0
        self.paint.dll.AU3_WinActive.return_value = 1
        
        res = self.paint.change_thickness(direction=direction, count=count)
        
        mock_check_existing_window.assert_called_once()
        self.paint.dll.AU3_Send.assert_has_calls(
            [call("{ALT}sz", 0)]
        )
        self.paint.dll.AU3_WinActive.assert_not_called()
        
        self.assertEqual(self.paint.dll.AU3_Send.call_count, 1)
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not send keys to change thickness bar")

    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_change_failed_window_inactive_while_changing_thickness(self, mock_sleep, mock_check_existing_window):
        direction = "up"
        count = 10
        self.paint.dll.AU3_Send.return_value = 1
        self.paint.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        
        res = self.paint.change_thickness(direction=direction, count=count)
        
        mock_check_existing_window.assert_called_once()
        self.paint.dll.AU3_Send.assert_has_calls(
            [call("{ALT}sz", 0)] + 
            [call("{UP}", 0)] * 5
        )
        self.paint.dll.AU3_WinActive.assert_has_calls([call(self.paint.window_info, "")] * 6)
        
        self.assertEqual(self.paint.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.paint.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 7)
        self.assertEqual(res, "Microsoft Paint window is inactive")

    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_change_failed_to_send_keys_while_changing_thickness_up(self, mock_sleep, mock_check_existing_window):
        direction = "up"
        count = 10
        self.paint.dll.AU3_Send.side_effect = [1] * 5 + [0]
        self.paint.dll.AU3_WinActive.return_value = 1
        
        res = self.paint.change_thickness(direction=direction, count=count)
        
        mock_check_existing_window.assert_called_once()
        self.paint.dll.AU3_Send.assert_has_calls(
            [call("{ALT}sz", 0)] + 
            [call("{UP}", 0)] * 5
        )
        self.paint.dll.AU3_WinActive.assert_has_calls([call(self.paint.window_info, "")] * 5)
        
        self.assertEqual(self.paint.dll.AU3_WinActive.call_count, 5)
        self.assertEqual(self.paint.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "could not send keys to change thickness")

    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_change_failed_to_send_keys_while_changing_thickness_down(self, mock_sleep, mock_check_existing_window):
        direction = "down"
        count = 10
        self.paint.dll.AU3_Send.side_effect = [1] * 5 + [0]
        self.paint.dll.AU3_WinActive.return_value = 1
        
        res = self.paint.change_thickness(direction=direction, count=count)
        
        mock_check_existing_window.assert_called_once()
        self.paint.dll.AU3_Send.assert_has_calls(
            [call("{ALT}sz", 0)] + 
            [call("{DN}", 0)] * 5
        )
        self.paint.dll.AU3_WinActive.assert_has_calls([call(self.paint.window_info, "")] * 5)
        
        self.assertEqual(self.paint.dll.AU3_WinActive.call_count, 5)
        self.assertEqual(self.paint.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "could not send keys to change thickness")
        
    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_change_image_size_success(self, mock_sleep, mock_check_existing_window):
        width = 1920
        height = 1080
        self.paint.dll.AU3_Send.return_value = 1
        self.paint.dll.AU3_WinExists.return_value = 1
        self.paint.dll.AU3_WinActivate.return_value = 1
        self.paint.dll.AU3_WinWaitActive.return_value = 1
        
        res = self.paint.change_image_size(width=width, height=height)
        
        mock_check_existing_window.assert_called_once()
        self.paint.dll.AU3_Send.assert_has_calls([
            call("^e", 0),
            call(str(width), 1),
            call("{TAB}", 0),
            call(str(height), 1),
            call("{ENTER}", 0)
        ])
        self.paint.dll.AU3_WinExists.assert_called_once_with(self.paint.image_properties_window_info, "")
        self.paint.dll.AU3_WinActivate.assert_called_once_with(self.paint.image_properties_window_info, "")
        self.paint.dll.AU3_WinWaitActive.assert_called_once_with(self.paint.image_properties_window_info, "", 10)
        
        self.assertEqual(self.paint.dll.AU3_Send.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertIsNone(res)
        
    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_change_image_size_failed_image_width_is_not_provided(self, mock_sleep, mock_check_existing_window):
        width = 0
        height = 1080
        
        res = self.paint.change_image_size(width=width, height=height)
        
        mock_check_existing_window.assert_not_called()
        self.paint.dll.AU3_Send.assert_not_called()
        self.paint.dll.AU3_WinExists.assert_not_called()
        self.paint.dll.AU3_WinActivate.assert_not_called()
        self.paint.dll.AU3_WinWaitActive.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, "image width must be provided")
        
    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_change_image_size_failed_image_height_is_not_provided(self, mock_sleep, mock_check_existing_window):
        width = 1920
        height = 0
        
        res = self.paint.change_image_size(width=width, height=height)
        
        mock_check_existing_window.assert_not_called()
        self.paint.dll.AU3_Send.assert_not_called()
        self.paint.dll.AU3_WinExists.assert_not_called()
        self.paint.dll.AU3_WinActivate.assert_not_called()
        self.paint.dll.AU3_WinWaitActive.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, "image height must be provided")
        
    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_change_image_size_failed_to_send_keys_to_activate_image_properties_modal(self, mock_sleep, mock_check_existing_window):
        width = 1920
        height = 1080
        self.paint.dll.AU3_Send.return_value = 0
        
        res = self.paint.change_image_size(width=width, height=height)
        
        mock_check_existing_window.assert_called_once()
        self.paint.dll.AU3_Send.assert_has_calls([
            call("^e", 0),
        ])
        self.paint.dll.AU3_WinExists.assert_not_called()
        self.paint.dll.AU3_WinActivate.assert_not_called()
        self.paint.dll.AU3_WinWaitActive.assert_not_called()
                
        self.assertEqual(self.paint.dll.AU3_Send.call_count, 1)
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not send keys to activate image properties modal")
        
    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_change_image_size_failed_image_properties_modal_window_didnt_exist(self, mock_sleep, mock_check_existing_window):
        width = 1920
        height = 1080
        self.paint.dll.AU3_Send.return_value = 1
        self.paint.dll.AU3_WinExists.return_value = 0
        
        res = self.paint.change_image_size(width=width, height=height)
        
        mock_check_existing_window.assert_called_once()
        self.paint.dll.AU3_Send.assert_has_calls([
            call("^e", 0),
        ])
        self.paint.dll.AU3_WinExists.assert_called_once_with(self.paint.image_properties_window_info, "")
        self.paint.dll.AU3_WinActivate.assert_not_called()
        self.paint.dll.AU3_WinWaitActive.assert_not_called()
                
        self.assertEqual(self.paint.dll.AU3_Send.call_count, 1)
        self.assertEqual(mock_sleep.call_count, 2)
        self.assertEqual(res, "image properties modal didn't exist")
        
    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_change_image_size_failed_to_wait_window_activation(self, mock_sleep, mock_check_existing_window):
        width = 1920
        height = 1080
        self.paint.dll.AU3_Send.return_value = 1
        self.paint.dll.AU3_WinExists.return_value = 1
        self.paint.dll.AU3_WinActivate.return_value = 1
        self.paint.dll.AU3_WinWaitActive.return_value = 0
        
        res = self.paint.change_image_size(width=width, height=height)
        
        mock_check_existing_window.assert_called_once()
        self.paint.dll.AU3_Send.assert_called_once_with("^e", 0)
        self.paint.dll.AU3_WinExists.assert_called_once_with(self.paint.image_properties_window_info, "")
        self.paint.dll.AU3_WinActivate.assert_called_once_with(self.paint.image_properties_window_info, "")
        self.paint.dll.AU3_WinWaitActive.assert_called_once_with(self.paint.image_properties_window_info, "", 10)
        
        self.assertEqual(self.paint.dll.AU3_Send.call_count, 1)
        self.assertEqual(mock_sleep.call_count, 2)
        self.assertEqual(res, "could not activate image properties modal")
        
    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_change_image_size_failed_to_send_image_width(self, mock_sleep, mock_check_existing_window):
        width = 1920
        height = 1080
        self.paint.dll.AU3_Send.side_effect = [1, 0]
        self.paint.dll.AU3_WinExists.return_value = 1
        self.paint.dll.AU3_WinActivate.return_value = 1
        self.paint.dll.AU3_WinWaitActive.return_value = 1
        
        res = self.paint.change_image_size(width=width, height=height)
        
        mock_check_existing_window.assert_called_once()
        self.paint.dll.AU3_Send.assert_has_calls([
            call("^e", 0),
            call(str(width), 1)
        ])
        self.paint.dll.AU3_WinExists.assert_called_once_with(self.paint.image_properties_window_info, "")
        self.paint.dll.AU3_WinActivate.assert_called_once_with(self.paint.image_properties_window_info, "")
        self.paint.dll.AU3_WinWaitActive.assert_called_once_with(self.paint.image_properties_window_info, "", 10)
        
        self.assertEqual(self.paint.dll.AU3_Send.call_count, 2)
        self.assertEqual(mock_sleep.call_count, 3)
        self.assertEqual(res, "could not send image width to image properties modal")
        
    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_change_image_size_failed_to_send_key_to_modal(self, mock_sleep, mock_check_existing_window):
        width = 1920
        height = 1080
        self.paint.dll.AU3_Send.side_effect = [1, 1, 0]
        self.paint.dll.AU3_WinExists.return_value = 1
        self.paint.dll.AU3_WinActivate.return_value = 1
        self.paint.dll.AU3_WinWaitActive.return_value = 1
        
        res = self.paint.change_image_size(width=width, height=height)
        
        mock_check_existing_window.assert_called_once()
        self.paint.dll.AU3_Send.assert_has_calls([
            call("^e", 0),
            call(str(width), 1),
            call("{TAB}", 0)
        ])
        self.paint.dll.AU3_WinExists.assert_called_once_with(self.paint.image_properties_window_info, "")
        self.paint.dll.AU3_WinActivate.assert_called_once_with(self.paint.image_properties_window_info, "")
        self.paint.dll.AU3_WinWaitActive.assert_called_once_with(self.paint.image_properties_window_info, "", 10)
        
        self.assertEqual(self.paint.dll.AU3_Send.call_count, 3)
        self.assertEqual(mock_sleep.call_count, 4)
        self.assertEqual(res, "could not send tab to image properties modal")
        
    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_change_image_size_failed_to_send_image_height(self, mock_sleep, mock_check_existing_window):
        width = 1920
        height = 1080
        self.paint.dll.AU3_Send.side_effect = [1, 1, 1, 0]
        self.paint.dll.AU3_WinExists.return_value = 1
        self.paint.dll.AU3_WinActivate.return_value = 1
        self.paint.dll.AU3_WinWaitActive.return_value = 1
        
        res = self.paint.change_image_size(width=width, height=height)
        
        mock_check_existing_window.assert_called_once()
        self.paint.dll.AU3_Send.assert_has_calls([
            call("^e", 0),
            call(str(width), 1),
            call("{TAB}", 0),
            call(str(height), 1)
        ])
        self.paint.dll.AU3_WinExists.assert_called_once_with(self.paint.image_properties_window_info, "")
        self.paint.dll.AU3_WinActivate.assert_called_once_with(self.paint.image_properties_window_info, "")
        self.paint.dll.AU3_WinWaitActive.assert_called_once_with(self.paint.image_properties_window_info, "", 10)
        
        self.assertEqual(self.paint.dll.AU3_Send.call_count, 4)
        self.assertEqual(mock_sleep.call_count, 5)
        self.assertEqual(res, "could not send image height to image properties modal")
        
    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_change_image_size_failed_to_send_key_to_change_image_size(self, mock_sleep, mock_check_existing_window):
        width = 1920
        height = 1080
        self.paint.dll.AU3_Send.side_effect = [1, 1, 1, 1, 0]
        self.paint.dll.AU3_WinExists.return_value = 1
        self.paint.dll.AU3_WinActivate.return_value = 1
        self.paint.dll.AU3_WinWaitActive.return_value = 1
        
        res = self.paint.change_image_size(width=width, height=height)
        
        mock_check_existing_window.assert_called_once()
        self.paint.dll.AU3_Send.assert_has_calls([
            call("^e", 0),
            call(str(width), 1),
            call("{TAB}", 0),
            call(str(height), 1),
            call("{ENTER}", 0)
        ])
        self.paint.dll.AU3_WinExists.assert_called_once_with(self.paint.image_properties_window_info, "")
        self.paint.dll.AU3_WinActivate.assert_called_once_with(self.paint.image_properties_window_info, "")
        self.paint.dll.AU3_WinWaitActive.assert_called_once_with(self.paint.image_properties_window_info, "", 10)
        
        self.assertEqual(self.paint.dll.AU3_Send.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "could not send Enter key to change image size")
        
    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None) 
    def test_save_file_success(self, mock_check_existing_window):
        self.paint.dll.AU3_Send.return_value = 1
        
        res = self.paint.save_file()
        
        mock_check_existing_window.assert_called_once() 
        self.paint.dll.AU3_Send.assert_called_once_with("^s", 0)
        self.assertIsNone(res)
        
    @patch.object(MicrosoftPaint, "check_existing_window", return_value=None) 
    def test_save_file_failed(self, mock_check_existing_window):
        self.paint.dll.AU3_Send.return_value = 0
        
        res = self.paint.save_file()
        
        mock_check_existing_window.assert_called_once() 
        self.paint.dll.AU3_Send.assert_called_once_with("^s", 0)
        self.assertEqual(res, "could not send ctrl+s to save file")