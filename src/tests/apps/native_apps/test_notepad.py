import unittest
from unittest.mock import patch, call, MagicMock
from activities.apps.native_apps.notepad import Notepad

class TestNotepad(unittest.TestCase):
    def setUp(self):
        self.notepad = Notepad()
        self.notepad.dll = MagicMock()

    def test_check_existing_window_success(self):
        self.notepad.dll.AU3_WinExists.return_value = 1
        self.notepad.dll.AU3_WinWaitActive.return_value = 1

        res = self.notepad.check_existing_window()

        self.notepad.dll.AU3_WinExists.assert_called_once_with(self.notepad.window_info, "")
        self.notepad.dll.AU3_WinActivate.assert_called_once_with(self.notepad.window_info, "")
        self.notepad.dll.AU3_WinWaitActive.assert_called_once_with(self.notepad.window_info, "", 10)
        
        self.assertIsNone(res)
        
    def test_check_existing_window_failed_window_didnt_exist(self):
        self.notepad.dll.AU3_WinExists.return_value = 0
        res = self.notepad.check_existing_window()
        
        self.notepad.dll.AU3_WinExists.assert_called_once_with(self.notepad.window_info, "")
        self.notepad.dll.AU3_WinActivate.assert_not_called()
        self.notepad.dll.AU3_WinWaitActive.assert_not_called()
        
        self.assertEqual(res, "Notepad window didn't exist")
        
    def test_check_existing_window_failed_to_wait_window_activation(self):
        self.notepad.dll.AU3_WinExists.return_value = 1
        self.notepad.dll.AU3_WinActivate.return_value = 1
        self.notepad.dll.AU3_WinWaitActive.return_value = 0
        
        res = self.notepad.check_existing_window()

        self.notepad.dll.AU3_WinExists.assert_called_once_with(self.notepad.window_info, "")
        self.notepad.dll.AU3_WinActivate.assert_called_once_with(self.notepad.window_info, "")
        self.notepad.dll.AU3_WinWaitActive.assert_called_once_with(self.notepad.window_info, "", 10)
        
        self.assertEqual(res, "could not activate Notepad window")

    @patch("time.sleep", return_value=None)
    @patch.object(Notepad, "_get_executable_path", return_value="/path/to/notepad/executable")
    def test_create_window_success(self, mock_get_executable_path, mock_sleep):
        self.notepad.dll.AU3_Run.return_value = 1
        self.notepad.dll.AU3_WinExists.return_value = 1
        self.notepad.dll.AU3_WinGetHandle.return_value = 1234567890
        self.notepad.dll.AU3_WinActivate.return_value = 1
        self.notepad.dll.AU3_WinWaitActive.return_value = 1
        self.notepad.dll.AU3_WinSetState.return_value = 1

        res = self.notepad.create_window()
        
        mock_get_executable_path.assert_called_once()
        self.notepad.dll.AU3_Run.assert_called_once_with("/path/to/notepad/executable", "", 1)
        self.notepad.dll.AU3_WinExists.assert_called_once()
        self.notepad.dll.AU3_WinGetHandle.assert_called_once()
        self.notepad.dll.AU3_WinActivate.assert_called_once_with(self.notepad.window_info, "")
        self.notepad.dll.AU3_WinWaitActive.assert_called_once_with(self.notepad.window_info, "", 10)
        self.notepad.dll.AU3_WinSetState.assert_called_once_with(self.notepad.window_info, "", 3)

        self.assertEqual(mock_sleep.call_count, 2)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    @patch.object(Notepad, "_get_executable_path", return_value=None)
    def test_create_window_failed_to_get_executable_path(self, mock_get_executable_path, mock_sleep):
        res = self.notepad.create_window()
        
        mock_get_executable_path.assert_called_once()
        self.notepad.dll.AU3_Run.assert_not_called()
        self.notepad.dll.AU3_WinExists.assert_not_called()
        self.notepad.dll.AU3_WinGetHandle.assert_not_called()
        self.notepad.dll.AU3_WinActivate.assert_not_called()
        self.notepad.dll.AU3_WinWaitActive.assert_not_called()
        self.notepad.dll.AU3_WinSetState.assert_not_called()

        self.assertEqual(res, "could not get Notepad executable path")
        
    @patch("time.sleep", return_value=None)
    @patch.object(Notepad, "_get_executable_path", return_value="/path/to/notepad/executable")
    def test_create_window_failed_to_run(self, mock_get_executable_path, mock_sleep):
        self.notepad.dll.AU3_Run.return_value = 0
        res = self.notepad.create_window()
        
        mock_get_executable_path.assert_called_once()
        self.notepad.dll.AU3_Run.assert_called_once_with("/path/to/notepad/executable", "", 1)
        self.notepad.dll.AU3_WinExists.assert_not_called()
        self.notepad.dll.AU3_WinGetHandle.assert_not_called()
        self.notepad.dll.AU3_WinActivate.assert_not_called()
        self.notepad.dll.AU3_WinWaitActive.assert_not_called()
        self.notepad.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(res, "could not run Notepad")
        
    @patch("time.sleep", return_value=None)
    @patch.object(Notepad, "_get_executable_path", return_value="/path/to/notepad/executable")
    def test_create_window_failed_window_didnt_exist(self, mock_get_executable_path, mock_sleep):
        self.notepad.dll.AU3_Run.return_value = 1
        self.notepad.dll.AU3_WinExists.return_value = 0

        res = self.notepad.create_window()
        
        mock_get_executable_path.assert_called_once()
        self.notepad.dll.AU3_Run.assert_called_once_with("/path/to/notepad/executable", "", 1)
        self.notepad.dll.AU3_WinExists.assert_called_once()
        self.notepad.dll.AU3_WinGetHandle.assert_not_called()
        self.notepad.dll.AU3_WinActivate.assert_not_called()
        self.notepad.dll.AU3_WinWaitActive.assert_not_called()
        self.notepad.dll.AU3_WinSetState.assert_not_called()

        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "Notepad window didn't exist")
        
    @patch("time.sleep", return_value=None)
    @patch.object(Notepad, "_get_executable_path", return_value="/path/to/notepad/executable")
    def test_create_window_failed_to_get_handle(self, mock_get_executable_path, mock_sleep):
        self.notepad.dll.AU3_Run.return_value = 1
        self.notepad.dll.AU3_WinExists.return_value = 1
        self.notepad.dll.AU3_WinGetHandle.return_value = 0

        res = self.notepad.create_window()
        
        mock_get_executable_path.assert_called_once()
        self.notepad.dll.AU3_Run.assert_called_once_with("/path/to/notepad/executable", "", 1)
        self.notepad.dll.AU3_WinExists.assert_called_once()
        self.notepad.dll.AU3_WinGetHandle.assert_called_once()
        self.notepad.dll.AU3_WinActivate.assert_not_called()
        self.notepad.dll.AU3_WinWaitActive.assert_not_called()
        self.notepad.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not get window handle")
        
    @patch("time.sleep", return_value=None)
    @patch.object(Notepad, "_get_executable_path", return_value="/path/to/notepad/executable")
    def test_create_window_failed_to_wait_window_activation(self, mock_get_executable_path, mock_sleep):
        self.notepad.dll.AU3_Run.return_value = 1
        self.notepad.dll.AU3_WinExists.return_value = 1
        self.notepad.dll.AU3_WinGetHandle.return_value = 1234567890
        self.notepad.dll.AU3_WinActivate.return_value = 1
        self.notepad.dll.AU3_WinWaitActive.return_value = 0

        res = self.notepad.create_window()
        
        mock_get_executable_path.assert_called_once()
        self.notepad.dll.AU3_Run.assert_called_once_with("/path/to/notepad/executable", "", 1)
        self.notepad.dll.AU3_WinExists.assert_called_once()
        self.notepad.dll.AU3_WinGetHandle.assert_called_once()
        self.notepad.dll.AU3_WinActivate.assert_called_once_with(self.notepad.window_info, "")
        self.notepad.dll.AU3_WinWaitActive.assert_called_once_with(self.notepad.window_info, "", 10)
        self.notepad.dll.AU3_WinSetState.assert_not_called()

        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not activate Notepad")
        
    @patch("time.sleep", return_value=None)
    @patch.object(Notepad, "_get_executable_path", return_value="/path/to/notepad/executable")
    def test_create_window_failed_to_maximize_window(self, mock_get_executable_path, mock_sleep):
        self.notepad.dll.AU3_Run.return_value = 1
        self.notepad.dll.AU3_WinExists.return_value = 1
        self.notepad.dll.AU3_WinGetHandle.return_value = 1234567890
        self.notepad.dll.AU3_WinActivate.return_value = 1
        self.notepad.dll.AU3_WinWaitActive.return_value = 1
        self.notepad.dll.AU3_WinSetState.return_value = 0

        res = self.notepad.create_window()
        
        mock_get_executable_path.assert_called_once()
        self.notepad.dll.AU3_Run.assert_called_once_with("/path/to/notepad/executable", "", 1)
        self.notepad.dll.AU3_WinExists.assert_called_once()
        self.notepad.dll.AU3_WinGetHandle.assert_called_once()
        self.notepad.dll.AU3_WinActivate.assert_called_once_with(self.notepad.window_info, "")
        self.notepad.dll.AU3_WinWaitActive.assert_called_once_with(self.notepad.window_info, "", 10)
        self.notepad.dll.AU3_WinSetState.assert_called_once_with(self.notepad.window_info, "", 3)

        self.assertEqual(mock_sleep.call_count, 2)
        self.assertEqual(res, "could not maximize Notepad window")
        
    @patch("time.sleep", return_value=None)
    @patch.object(Notepad, "check_existing_window", return_value=None)
    def test_create_tab_success(self, mock_check_existing_window, mock_sleep):
        self.notepad.dll.AU3_WinSetState.return_value = 1
        self.notepad.dll.AU3_Send.return_value = 1
    
        res = self.notepad.create_tab()
        
        mock_check_existing_window.assert_called_once()
        self.notepad.dll.AU3_WinSetState.assert_called_once_with(self.notepad.window_info, "", 3)
        self.notepad.dll.AU3_Send.assert_called_once_with("^t", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    @patch.object(Notepad, "check_existing_window", return_value=None)
    def test_create_tab_failed_to_maximize_window(self, mock_check_existing_window, mock_sleep):
        self.notepad.dll.AU3_WinSetState.return_value = 0
    
        res = self.notepad.create_tab()
        
        mock_check_existing_window.assert_called_once()
        self.notepad.dll.AU3_WinSetState.assert_called_once_with(self.notepad.window_info, "", 3)
        self.notepad.dll.AU3_Send.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not maximize Notepad window")
        
    @patch("time.sleep", return_value=None)
    @patch.object(Notepad, "check_existing_window", return_value=None)
    def test_create_tab_failed_to_send_keys(self, mock_check_existing_window, mock_sleep):
        self.notepad.dll.AU3_WinSetState.return_value = 1
        self.notepad.dll.AU3_Send.return_value = 0
    
        res = self.notepad.create_tab()
        
        mock_check_existing_window.assert_called_once()
        self.notepad.dll.AU3_WinSetState.assert_called_once_with(self.notepad.window_info, "", 3)
        self.notepad.dll.AU3_Send.assert_called_once_with("^t", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not send keys to create new tab")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(Notepad, "_get_executable_path", return_value="/path/to/notepad/executable")
    def test_open_file_success(self, mock_get_executable_path, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        self.notepad.dll.AU3_Run.return_value = 1
        self.notepad.dll.AU3_WinExists.return_value = 1
        self.notepad.dll.AU3_WinGetHandle.return_value = 1234567890
        self.notepad.dll.AU3_WinActivate.return_value = 1
        self.notepad.dll.AU3_WinWaitActive.return_value = 1
        self.notepad.dll.AU3_WinSetState.return_value = 1
        
        res = self.notepad.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_get_executable_path.assert_called_once()
        self.notepad.dll.AU3_Run.assert_called_once_with(f'/path/to/notepad/executable "{path}"', "", 1)
        self.notepad.dll.AU3_WinExists.assert_called_once()
        self.notepad.dll.AU3_WinGetHandle.assert_called_once()
        self.notepad.dll.AU3_WinActivate.assert_called_once_with(self.notepad.window_info, "")
        self.notepad.dll.AU3_WinWaitActive.assert_called_once_with(self.notepad.window_info, "", 10)
        self.notepad.dll.AU3_WinSetState.assert_called_once_with(self.notepad.window_info, "", 3)
        
        self.assertEqual(mock_sleep.call_count, 2)
        self.assertIsNone(res)

    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(Notepad, "_get_executable_path")
    def test_open_file_failed_path_is_not_provided(self, mock_get_executable_path, mock_os_path, mock_sleep):
        res = self.notepad.open_file(path="")
        
        mock_os_path.exists.assert_not_called()
        mock_os_path.isfile.assert_not_called()
        mock_get_executable_path.assert_not_called()
        self.notepad.dll.AU3_Run.assert_not_called()
        self.notepad.dll.AU3_WinExists.assert_not_called()
        self.notepad.dll.AU3_WinGetHandle.assert_not_called()
        self.notepad.dll.AU3_WinActivate.assert_not_called()
        self.notepad.dll.AU3_WinWaitActive.assert_not_called()
        self.notepad.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 0)
        self.assertEqual(res, "path must be provided")

    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(Notepad, "_get_executable_path")  
    def test_open_file_failed_path_didnt_exist(self, mock_get_executable_path, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = False
        
        res = self.notepad.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_not_called()
        mock_get_executable_path.assert_not_called()
        self.notepad.dll.AU3_Run.assert_not_called()
        self.notepad.dll.AU3_WinExists.assert_not_called()
        self.notepad.dll.AU3_WinGetHandle.assert_not_called()
        self.notepad.dll.AU3_WinActivate.assert_not_called()
        self.notepad.dll.AU3_WinWaitActive.assert_not_called()
        self.notepad.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 0)
        self.assertEqual(res, f"file path '{path}' does not exist")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(Notepad, "_get_executable_path")
    def test_open_file_failed_path_is_not_a_file(self, mock_get_executable_path, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = False
        
        res = self.notepad.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_get_executable_path.assert_not_called()
        self.notepad.dll.AU3_Run.assert_not_called()
        self.notepad.dll.AU3_WinExists.assert_not_called()
        self.notepad.dll.AU3_WinGetHandle.assert_not_called()
        self.notepad.dll.AU3_WinActivate.assert_not_called()
        self.notepad.dll.AU3_WinWaitActive.assert_not_called()
        self.notepad.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 0)
        self.assertEqual(res, f"path '{path}' is not a file")
               
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(Notepad, "_get_executable_path", return_value=None)
    def test_open_file_failed_to_get_executable_path(self, mock_get_executable_path, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        
        res = self.notepad.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_get_executable_path.asset_called_once()
        self.notepad.dll.AU3_Run.assert_not_called()
        self.notepad.dll.AU3_WinExists.assert_not_called()
        self.notepad.dll.AU3_WinGetHandle.assert_not_called()
        self.notepad.dll.AU3_WinActivate.assert_not_called()
        self.notepad.dll.AU3_WinWaitActive.assert_not_called()
        self.notepad.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 0)
        self.assertEqual(res, "could not get Notepad executable path")
          
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(Notepad, "_get_executable_path", return_value="/path/to/notepad/executable")
    def test_open_file_failed_to_run(self, mock_get_executable_path, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        self.notepad.dll.AU3_Run.return_value = 0
        
        res = self.notepad.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_get_executable_path.assert_called_once()
        self.notepad.dll.AU3_Run.assert_called_once_with(f'/path/to/notepad/executable "{path}"', "", 1)
        self.notepad.dll.AU3_WinExists.assert_not_called()
        self.notepad.dll.AU3_WinGetHandle.assert_not_called()
        self.notepad.dll.AU3_WinActivate.assert_not_called()
        self.notepad.dll.AU3_WinWaitActive.assert_not_called()
        self.notepad.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 0)
        self.assertEqual(res, "could not open file with notepad")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(Notepad, "_get_executable_path", return_value="/path/to/notepad/executable")
    def test_open_file_failed_window_didnt_exist(self, mock_get_executable_path, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        self.notepad.dll.AU3_Run.return_value = 1
        self.notepad.dll.AU3_WinExists.return_value = 0
        
        res = self.notepad.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_get_executable_path.assert_called_once()
        self.notepad.dll.AU3_Run.assert_called_once_with(f'/path/to/notepad/executable "{path}"', "", 1)
        self.notepad.dll.AU3_WinExists.assert_called_once()
        self.notepad.dll.AU3_WinGetHandle.assert_not_called()
        self.notepad.dll.AU3_WinActivate.assert_not_called()
        self.notepad.dll.AU3_WinWaitActive.assert_not_called()
        self.notepad.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "Notepad window didn't exist")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(Notepad, "_get_executable_path", return_value="/path/to/notepad/executable")
    def test_open_file_failed_to_get_handle(self, mock_get_executable_path, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        self.notepad.dll.AU3_Run.return_value = 1
        self.notepad.dll.AU3_WinExists.return_value = 1
        self.notepad.dll.AU3_WinGetHandle.return_value = 0
        
        res = self.notepad.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_get_executable_path.assert_called_once()
        self.notepad.dll.AU3_Run.assert_called_once_with(f'/path/to/notepad/executable "{path}"', "", 1)
        self.notepad.dll.AU3_WinExists.assert_called_once()
        self.notepad.dll.AU3_WinGetHandle.assert_called_once()
        self.notepad.dll.AU3_WinActivate.assert_not_called()
        self.notepad.dll.AU3_WinWaitActive.assert_not_called()
        self.notepad.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not get window handle")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(Notepad, "_get_executable_path", return_value="/path/to/notepad/executable")
    def test_open_file_failed_to_wait_window_activation(self, mock_get_executable_path, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        self.notepad.dll.AU3_Run.return_value = 1
        self.notepad.dll.AU3_WinExists.return_value = 1
        self.notepad.dll.AU3_WinGetHandle.return_value = 1234567890
        self.notepad.dll.AU3_WinActivate.return_value = 1
        self.notepad.dll.AU3_WinWaitActive.return_value = 0
        
        res = self.notepad.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_get_executable_path.assert_called_once()
        self.notepad.dll.AU3_Run.assert_called_once_with(f'/path/to/notepad/executable "{path}"', "", 1)
        self.notepad.dll.AU3_WinExists.assert_called_once()
        self.notepad.dll.AU3_WinGetHandle.assert_called_once()
        self.notepad.dll.AU3_WinActivate.assert_called_once_with(self.notepad.window_info, "")
        self.notepad.dll.AU3_WinWaitActive.assert_called_once_with(self.notepad.window_info, "", 10)
        self.notepad.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not activate Notepad")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(Notepad, "_get_executable_path", return_value="/path/to/notepad/executable")
    def test_open_file_failed_to_maximize_window(self, mock_get_executable_path, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        self.notepad.dll.AU3_Run.return_value = 1
        self.notepad.dll.AU3_WinExists.return_value = 1
        self.notepad.dll.AU3_WinGetHandle.return_value = 1234567890
        self.notepad.dll.AU3_WinActivate.return_value = 1
        self.notepad.dll.AU3_WinWaitActive.return_value = 1
        self.notepad.dll.AU3_WinSetState.return_value = 0
        
        res = self.notepad.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_get_executable_path.assert_called_once()
        self.notepad.dll.AU3_Run.assert_called_once_with(f'/path/to/notepad/executable "{path}"', "", 1)
        self.notepad.dll.AU3_WinExists.assert_called_once()
        self.notepad.dll.AU3_WinGetHandle.assert_called_once()
        self.notepad.dll.AU3_WinActivate.assert_called_once_with(self.notepad.window_info, "")
        self.notepad.dll.AU3_WinWaitActive.assert_called_once_with(self.notepad.window_info, "", 10)
        self.notepad.dll.AU3_WinSetState.assert_called_once_with(self.notepad.window_info, "", 3)
        
        self.assertEqual(mock_sleep.call_count, 2)
        self.assertEqual(res, "could not maximize Notepad window")

    @patch.object(Notepad, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_write_file_success(self, mock_sleep, mock_check_existing_window):
        text = "lorem ipsum dolor sit amet"
        self.notepad.dll.AU3_Send.return_value = 1
        self.notepad.dll.AU3_WinActive.return_value = 1
        
        res = self.notepad.write_file(text=text)
        
        mock_check_existing_window.assert_called_once()
        self.notepad.dll.AU3_Send.assert_has_calls([
            call("^{END}", 0),
            *[call(char, 1) for char in text],
        ])
        self.notepad.dll.AU3_WinActive.assert_has_calls([call(self.notepad.window_info, "")] * len(text))
        
        self.assertEqual(self.notepad.dll.AU3_WinActive.call_count, len(text))
        self.assertEqual(self.notepad.dll.AU3_Send.call_count, len(text) + 1)
        self.assertEqual(mock_sleep.call_count, len(text) + 1)
        self.assertIsNone(res)

    @patch.object(Notepad, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_write_file_failed_to_send_keys_to_move_cursor(self, mock_sleep, mock_check_existing_window):
        text = "lorem ipsum dolor sit amet"
        self.notepad.dll.AU3_Send.return_value = 0
        self.notepad.dll.AU3_WinActive.return_value = 1
        
        res = self.notepad.write_file(text=text)
        
        mock_check_existing_window.assert_called_once()
        self.notepad.dll.AU3_Send.assert_has_calls([
            call("^{END}", 0)
        ])
        self.notepad.dll.AU3_WinActive.assert_not_called()
        mock_sleep.assert_not_called()
        
        self.assertEqual(self.notepad.dll.AU3_Send.call_count, 1)
        self.assertEqual(res, "could not send keys to move cursor")

    @patch.object(Notepad, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_write_file_failed_window_inactive_while_writing_file(self, mock_sleep, mock_check_existing_window):
        text = "lorem ipsum dolor sit amet"
        self.notepad.dll.AU3_Send.return_value = 1
        self.notepad.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        
        res = self.notepad.write_file(text=text)
        
        mock_check_existing_window.assert_called_once()
        self.notepad.dll.AU3_Send.assert_has_calls([
            call("^{END}", 0),
            *[call(char, 1) for char in "lorem"],
        ])
        self.notepad.dll.AU3_WinActive.assert_has_calls([call(self.notepad.window_info, "")] * 6)
        
        self.assertEqual(self.notepad.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.notepad.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "Notepad window is inactive")

    @patch.object(Notepad, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_write_file_failed_to_send_letter_while_writing_file(self, mock_sleep, mock_check_existing_window):
        text = "lorem ipsum dolor sit amet"
        self.notepad.dll.AU3_Send.side_effect = [1] * 5 + [0]
        self.notepad.dll.AU3_WinActive.return_value = 1
        
        res = self.notepad.write_file(text=text)
        
        mock_check_existing_window.assert_called_once()
        self.notepad.dll.AU3_Send.assert_has_calls([
            call("^{END}", 0),
            *[call(char, 1) for char in "lorem"],
        ])
        self.notepad.dll.AU3_WinActive.assert_has_calls([call(self.notepad.window_info, "")] * 5)
        
        self.assertEqual(self.notepad.dll.AU3_WinActive.call_count, 5)
        self.assertEqual(self.notepad.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 5)
        self.assertEqual(res, "could not send m to Notepad")

    @patch.object(Notepad, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_scroll_success(self, mock_sleep, mock_check_existing_window):
        direction = "down"
        clicks = 30
        scroll_delay = 0.07
        self.notepad.dll.AU3_WinActive.return_value = 1
        self.notepad.dll.AU3_MouseWheel.return_value = 1
        
        res = self.notepad.scroll(direction=direction, clicks=clicks, scroll_delay=scroll_delay)
        
        mock_check_existing_window.assert_called_once()
        self.notepad.dll.AU3_WinActive.assert_has_calls([call(self.notepad.window_info, "")] * clicks)
        self.notepad.dll.AU3_MouseWheel.assert_has_calls([call(direction, 1)] * clicks)
        
        self.assertEqual(self.notepad.dll.AU3_WinActive.call_count, clicks)
        self.assertEqual(self.notepad.dll.AU3_MouseWheel.call_count, clicks)
        self.assertEqual(mock_sleep.call_count, clicks + 1)
        self.assertIsNone(res)

    @patch.object(Notepad, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_scroll_failed_invalid_scroll_direction(self, mock_sleep, mock_check_existing_window):
        direction = "abc"
        clicks = 30
        scroll_delay = 0.07
        
        res = self.notepad.scroll(direction=direction, clicks=clicks, scroll_delay=scroll_delay)
        
        mock_check_existing_window.assert_not_called()
        self.notepad.dll.AU3_WinActive.assert_not_called()
        self.notepad.dll.AU3_MouseWheel.assert_not_called()
        mock_sleep.assert_not_called()

        self.assertEqual(res, "invalid scroll direction")

    @patch.object(Notepad, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_scroll_failed_window_inactive_while_scrolling(self, mock_sleep, mock_check_existing_window):
        direction = "down"
        clicks = 30
        scroll_delay = 0.07
        self.notepad.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        self.notepad.dll.AU3_MouseWheel.return_value = 1
        
        res = self.notepad.scroll(direction=direction, clicks=clicks, scroll_delay=scroll_delay)
        
        mock_check_existing_window.assert_called_once()
        self.notepad.dll.AU3_WinActive.assert_has_calls([call(self.notepad.window_info, "")] * 6)
        self.notepad.dll.AU3_MouseWheel.assert_has_calls([call(direction, 1)] * 5)
        
        self.assertEqual(self.notepad.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.notepad.dll.AU3_MouseWheel.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "Notepad window is inactive")

    @patch.object(Notepad, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_scroll_failed_to_scroll_mouse_wheel(self, mock_sleep, mock_check_existing_window):
        direction = "down"
        clicks = 30
        scroll_delay = 0.07
        self.notepad.dll.AU3_WinActive.return_value = 1
        self.notepad.dll.AU3_MouseWheel.side_effect = [1] * 5 + [0]
        
        res = self.notepad.scroll(direction=direction, clicks=clicks, scroll_delay=scroll_delay)
        
        mock_check_existing_window.assert_called_once()
        self.notepad.dll.AU3_WinActive.assert_has_calls([call(self.notepad.window_info, "")] * 6)
        self.notepad.dll.AU3_MouseWheel.assert_has_calls([call(direction, 1)] * 6)
        
        self.assertEqual(self.notepad.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.notepad.dll.AU3_MouseWheel.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 7)
        self.assertEqual(res, "could not scroll mouse wheel")
        
    @patch.object(Notepad, "check_existing_window", return_value=None) 
    def test_save_file_success(self, mock_check_existing_window):
        self.notepad.dll.AU3_Send.return_value = 1
        
        res = self.notepad.save_file()
        
        mock_check_existing_window.assert_called_once() 
        self.notepad.dll.AU3_Send.assert_called_once_with("^s", 0)
        self.assertIsNone(res)
        
    @patch.object(Notepad, "check_existing_window", return_value=None) 
    def test_save_file_failed(self, mock_check_existing_window):
        self.notepad.dll.AU3_Send.return_value = 0
        
        res = self.notepad.save_file()
        
        mock_check_existing_window.assert_called_once() 
        self.notepad.dll.AU3_Send.assert_called_once_with("^s", 0)
        self.assertEqual(res, "could not send ctrl+s to save file")