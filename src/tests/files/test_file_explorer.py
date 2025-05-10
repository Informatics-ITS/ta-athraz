import unittest
from unittest.mock import patch, call, MagicMock
from activities.files.file_explorer import FileExplorer

class TestFileExplorer(unittest.TestCase):
    def setUp(self):
        self.explorer = FileExplorer()
        self.explorer.dll = MagicMock()
        
    def test_check_existing_window_success(self):
        self.explorer.dll.AU3_WinExists.return_value = 1
        self.explorer.dll.AU3_WinWaitActive.return_value = 1

        res = self.explorer.check_existing_window()

        self.explorer.dll.AU3_WinExists.assert_called_once_with(self.explorer.window_info, "")
        self.explorer.dll.AU3_WinActivate.assert_called_once_with(self.explorer.window_info, "")
        self.explorer.dll.AU3_WinWaitActive.assert_called_once_with(self.explorer.window_info, "", 10)
        
        self.assertIsNone(res)
        
    def test_check_existing_window_failed_window_didnt_exist(self):
        self.explorer.dll.AU3_WinExists.return_value = 0
        res = self.explorer.check_existing_window()
        
        self.explorer.dll.AU3_WinExists.assert_called_once_with(self.explorer.window_info, "")
        self.explorer.dll.AU3_WinActivate.assert_not_called()
        self.explorer.dll.AU3_WinWaitActive.assert_not_called()
        
        self.assertEqual(res, "File Explorer window didn't exist")
        
    def test_check_existing_window_failed_to_wait_window_activation(self):
        self.explorer.dll.AU3_WinExists.return_value = 1
        self.explorer.dll.AU3_WinActivate.return_value = 1
        self.explorer.dll.AU3_WinWaitActive.return_value = 0
        
        res = self.explorer.check_existing_window()

        self.explorer.dll.AU3_WinExists.assert_called_once_with(self.explorer.window_info, "")
        self.explorer.dll.AU3_WinActivate.assert_called_once_with(self.explorer.window_info, "")
        self.explorer.dll.AU3_WinWaitActive.assert_called_once_with(self.explorer.window_info, "", 10)
        
        self.assertEqual(res, "could not activate File Explorer window")
        
    @patch("time.sleep", return_value=None)
    def test_create_window_success(self, mock_sleep):
        self.explorer.dll.AU3_Run.return_value = 1
        self.explorer.dll.AU3_WinExists.return_value = 1
        self.explorer.dll.AU3_WinGetHandle.return_value = 1234567890
        self.explorer.dll.AU3_WinActivate.return_value = 1
        self.explorer.dll.AU3_WinWaitActive.return_value = 1
        self.explorer.dll.AU3_WinSetState.return_value = 1

        res = self.explorer.create_window()

        self.explorer.dll.AU3_Run.assert_called_once_with("explorer.exe", "", 1)
        self.explorer.dll.AU3_WinExists.assert_called_once()
        self.explorer.dll.AU3_WinGetHandle.assert_called_once()
        self.explorer.dll.AU3_WinActivate.assert_called_once_with(self.explorer.window_info, "")
        self.explorer.dll.AU3_WinWaitActive.assert_called_once_with(self.explorer.window_info, "", 10)
        self.explorer.dll.AU3_WinSetState.assert_called_once_with(self.explorer.window_info, "", 3)

        self.assertEqual(mock_sleep.call_count, 2)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    def test_create_window_failed_to_run(self, mock_sleep):
        self.explorer.dll.AU3_Run.return_value = 0
        res = self.explorer.create_window()

        self.explorer.dll.AU3_Run.assert_called_once_with("explorer.exe", "", 1)
        self.explorer.dll.AU3_WinExists.assert_not_called()
        self.explorer.dll.AU3_WinGetHandle.assert_not_called()
        self.explorer.dll.AU3_WinActivate.assert_not_called()
        self.explorer.dll.AU3_WinWaitActive.assert_not_called()
        self.explorer.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(res, "could not run File Explorer")
        
    @patch("time.sleep", return_value=None)
    def test_create_window_failed_window_didnt_exist(self, mock_sleep):
        self.explorer.dll.AU3_Run.return_value = 1
        self.explorer.dll.AU3_WinExists.return_value = 0

        res = self.explorer.create_window()

        self.explorer.dll.AU3_Run.assert_called_once_with("explorer.exe", "", 1)
        self.explorer.dll.AU3_WinExists.assert_called_once()
        self.explorer.dll.AU3_WinGetHandle.assert_not_called()
        self.explorer.dll.AU3_WinActivate.assert_not_called()
        self.explorer.dll.AU3_WinWaitActive.assert_not_called()
        self.explorer.dll.AU3_WinSetState.assert_not_called()

        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "File Explorer window didn't exist")
        
    @patch("time.sleep", return_value=None)
    def test_create_window_failed_to_get_handle(self, mock_sleep):
        self.explorer.dll.AU3_Run.return_value = 1
        self.explorer.dll.AU3_WinExists.return_value = 1
        self.explorer.dll.AU3_WinGetHandle.return_value = 0

        res = self.explorer.create_window()

        self.explorer.dll.AU3_Run.assert_called_once_with("explorer.exe", "", 1)
        self.explorer.dll.AU3_WinExists.assert_called_once()
        self.explorer.dll.AU3_WinGetHandle.assert_called_once()
        self.explorer.dll.AU3_WinActivate.assert_not_called()
        self.explorer.dll.AU3_WinWaitActive.assert_not_called()
        self.explorer.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not get window handle")
        
    @patch("time.sleep", return_value=None)
    def test_create_window_failed_to_wait_window_activation(self, mock_sleep):
        self.explorer.dll.AU3_Run.return_value = 1
        self.explorer.dll.AU3_WinExists.return_value = 1
        self.explorer.dll.AU3_WinGetHandle.return_value = 1234567890
        self.explorer.dll.AU3_WinActivate.return_value = 1
        self.explorer.dll.AU3_WinWaitActive.return_value = 0

        res = self.explorer.create_window()

        self.explorer.dll.AU3_Run.assert_called_once_with("explorer.exe", "", 1)
        self.explorer.dll.AU3_WinExists.assert_called_once()
        self.explorer.dll.AU3_WinGetHandle.assert_called_once()
        self.explorer.dll.AU3_WinActivate.assert_called_once_with(self.explorer.window_info, "")
        self.explorer.dll.AU3_WinWaitActive.assert_called_once_with(self.explorer.window_info, "", 10)
        self.explorer.dll.AU3_WinSetState.assert_not_called()

        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not activate File Explorer")
        
    @patch("time.sleep", return_value=None)
    def test_create_window_failed_to_maximize_window(self, mock_sleep):
        self.explorer.dll.AU3_Run.return_value = 1
        self.explorer.dll.AU3_WinExists.return_value = 1
        self.explorer.dll.AU3_WinGetHandle.return_value = 1234567890
        self.explorer.dll.AU3_WinActivate.return_value = 1
        self.explorer.dll.AU3_WinWaitActive.return_value = 1
        self.explorer.dll.AU3_WinSetState.return_value = 0

        res = self.explorer.create_window()

        self.explorer.dll.AU3_Run.assert_called_once_with("explorer.exe", "", 1)
        self.explorer.dll.AU3_WinExists.assert_called_once()
        self.explorer.dll.AU3_WinGetHandle.assert_called_once()
        self.explorer.dll.AU3_WinActivate.assert_called_once_with(self.explorer.window_info, "")
        self.explorer.dll.AU3_WinWaitActive.assert_called_once_with(self.explorer.window_info, "", 10)
        self.explorer.dll.AU3_WinSetState.assert_called_once_with(self.explorer.window_info, "", 3)

        self.assertEqual(mock_sleep.call_count, 2)
        self.assertEqual(res, "could not maximize File Explorer window")
        
    @patch("time.sleep", return_value=None)
    @patch.object(FileExplorer, "check_existing_window", return_value=None)
    def test_create_tab_success(self, mock_check_existing_window, mock_sleep):
        self.explorer.dll.AU3_WinSetState.return_value = 1
        self.explorer.dll.AU3_Send.return_value = 1
        
        res = self.explorer.create_tab()
        
        mock_check_existing_window.assert_called_once()
        self.explorer.dll.AU3_WinSetState.assert_called_once_with(self.explorer.window_info, "", 3)
        self.explorer.dll.AU3_Send.assert_called_once_with("^t", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    @patch.object(FileExplorer, "check_existing_window", return_value=None)
    def test_create_tab_failed_to_maximize_window(self, mock_check_existing_window, mock_sleep):
        self.explorer.dll.AU3_WinSetState.return_value = 0
        
        res = self.explorer.create_tab()
        
        mock_check_existing_window.assert_called_once()
        self.explorer.dll.AU3_WinSetState.assert_called_once_with(self.explorer.window_info, "", 3)
        self.explorer.dll.AU3_Send.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not maximize File Explorer window")

    @patch("time.sleep", return_value=None)
    @patch.object(FileExplorer, "check_existing_window", return_value=None)
    def test_create_tab_failed_to_send_keys(self, mock_check_existing_window, mock_sleep):
        self.explorer.dll.AU3_WinSetState.return_value = 1
        self.explorer.dll.AU3_Send.return_value = 0
        
        res = self.explorer.create_tab()
        
        mock_check_existing_window.assert_called_once()
        self.explorer.dll.AU3_WinSetState.assert_called_once_with(self.explorer.window_info, "", 3)
        self.explorer.dll.AU3_Send.assert_called_once_with("^t", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not send keys to create new tab")

    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(FileExplorer, "check_existing_window", return_value=None)
    def test_open_file_success(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        self.explorer.dll.AU3_Send.return_value = 1
        self.explorer.dll.AU3_WinActive.return_value = 1
        
        res = self.explorer.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_check_existing_window.assert_called_once()
        self.explorer.dll.AU3_WinActive.assert_has_calls([call(self.explorer.window_info, "")] * len(path))
        self.explorer.dll.AU3_Send.assert_has_calls([
            call("^l", 0),
            *[call(char, 1) for char in path],
            call("{ENTER}", 0)
        ])
        
        self.assertEqual(self.explorer.dll.AU3_WinActive.call_count, len(path))
        self.assertEqual(self.explorer.dll.AU3_Send.call_count, len(path) + 2)
        self.assertEqual(mock_sleep.call_count, len(path) + 3)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(FileExplorer, "check_existing_window", return_value=None)
    def test_open_file_failed_path_is_not_provided(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = ""
        
        res = self.explorer.open_file(path=path)
        
        mock_os_path.exists.assert_not_called()
        mock_os_path.isfile.assert_not_called()
        mock_check_existing_window.assert_not_called()
        self.explorer.dll.AU3_WinActive.assert_not_called()
        self.explorer.dll.AU3_Send.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, "file path must be provided")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(FileExplorer, "check_existing_window", return_value=None)
    def test_open_file_failed_path_doesnt_exist(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = False
        
        res = self.explorer.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_not_called()
        mock_check_existing_window.assert_not_called()
        self.explorer.dll.AU3_WinActive.assert_not_called()
        self.explorer.dll.AU3_Send.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, f"file path '{path}' does not exist")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(FileExplorer, "check_existing_window", return_value=None)
    def test_open_file_failed_path_is_not_a_file(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = False
        self.explorer.dll.AU3_Send.return_value = 1
        self.explorer.dll.AU3_WinActive.return_value = 1
        
        res = self.explorer.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_check_existing_window.assert_not_called()
        self.explorer.dll.AU3_WinActive.assert_not_called()
        self.explorer.dll.AU3_Send.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, f"path '{path}' is not a file")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(FileExplorer, "check_existing_window", return_value=None)
    def test_open_file_failed_to_send_keys_to_open_address_bar(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        self.explorer.dll.AU3_Send.return_value = 0
        self.explorer.dll.AU3_WinActive.return_value = 1
        
        res = self.explorer.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_check_existing_window.assert_called_once()
        self.explorer.dll.AU3_WinActive.assert_not_called()
        self.explorer.dll.AU3_Send.assert_has_calls([
            call("^l", 0)
        ])
        
        self.assertEqual(self.explorer.dll.AU3_Send.call_count, 1)
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not send keys to open address bar")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(FileExplorer, "check_existing_window", return_value=None)
    def test_open_file_failed_window_inactive_while_filling_address_bar(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        self.explorer.dll.AU3_Send.return_value = 1
        self.explorer.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        
        res = self.explorer.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_check_existing_window.assert_called_once()
        self.explorer.dll.AU3_WinActive.assert_has_calls([call(self.explorer.window_info, "")] * 6)
        self.explorer.dll.AU3_Send.assert_has_calls([
            call("^l", 0),
            *[call(char, 1) for char in "path/"]
        ])
        
        self.assertEqual(self.explorer.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.explorer.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 7)
        self.assertEqual(res, "File Explorer window is inactive")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(FileExplorer, "check_existing_window", return_value=None)
    def test_open_file_failed_to_send_letter_while_filling_address_bar(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        self.explorer.dll.AU3_Send.side_effect = [1] * 5 + [0]
        self.explorer.dll.AU3_WinActive.return_value = 1
        
        res = self.explorer.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_check_existing_window.assert_called_once()
        self.explorer.dll.AU3_WinActive.assert_has_calls([call(self.explorer.window_info, "")] * 5)
        self.explorer.dll.AU3_Send.assert_has_calls([
            call("^l", 0),
            *[call(char, 1) for char in "path/"]
        ])
        
        self.assertEqual(self.explorer.dll.AU3_WinActive.call_count, 5)
        self.assertEqual(self.explorer.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "could not send / to File Explorer")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(FileExplorer, "check_existing_window", return_value=None)
    def test_open_file_failed_to_send_key_to_open_file(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        self.explorer.dll.AU3_Send.side_effect = [1] * len(path) + [1, 0]
        self.explorer.dll.AU3_WinActive.return_value = 1
        
        res = self.explorer.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_check_existing_window.assert_called_once()
        self.explorer.dll.AU3_WinActive.assert_has_calls([call(self.explorer.window_info, "")] * len(path))
        self.explorer.dll.AU3_Send.assert_has_calls([
            call("^l", 0),
            *[call(char, 1) for char in path],
            call("{ENTER}", 0)
        ])
        
        self.assertEqual(self.explorer.dll.AU3_WinActive.call_count, len(path))
        self.assertEqual(self.explorer.dll.AU3_Send.call_count, len(path) + 2)
        self.assertEqual(mock_sleep.call_count, len(path) + 3)
        self.assertEqual(res, "could not send enter key to open file")

    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(FileExplorer, "check_existing_window", return_value=None)
    def test_change_directory_success(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/to/directory"
        mock_os_path.exists.return_value = True
        mock_os_path.isdir.return_value = True
        self.explorer.dll.AU3_Send.return_value = 1
        self.explorer.dll.AU3_WinActive.return_value = 1
        
        res = self.explorer.change_directory(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isdir.assert_called_once_with(path)
        mock_check_existing_window.assert_called_once()
        self.explorer.dll.AU3_WinActive.assert_has_calls([call(self.explorer.window_info, "")] * len(path))
        self.explorer.dll.AU3_Send.assert_has_calls([
            call("^l", 0),
            *[call(char, 1) for char in path],
            call("{ENTER}", 0)
        ])
        
        self.assertEqual(self.explorer.dll.AU3_WinActive.call_count, len(path))
        self.assertEqual(self.explorer.dll.AU3_Send.call_count, len(path) + 2)
        self.assertEqual(mock_sleep.call_count, len(path) + 3)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(FileExplorer, "check_existing_window", return_value=None)
    def test_change_directory_failed_path_is_not_provided(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = ""
        
        res = self.explorer.change_directory(path=path)
        
        mock_os_path.exists.assert_not_called()
        mock_os_path.isdir.assert_not_called()
        mock_check_existing_window.assert_not_called()
        self.explorer.dll.AU3_WinActive.assert_not_called()
        self.explorer.dll.AU3_Send.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, "directory path must be provided")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(FileExplorer, "check_existing_window", return_value=None)
    def test_change_directory_failed_path_doesnt_exist(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/to/directory"
        mock_os_path.exists.return_value = False
        
        res = self.explorer.change_directory(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isdir.assert_not_called()
        mock_check_existing_window.assert_not_called()
        self.explorer.dll.AU3_WinActive.assert_not_called()
        self.explorer.dll.AU3_Send.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, f"directory '{path}' does not exist")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(FileExplorer, "check_existing_window", return_value=None)
    def test_change_directory_failed_path_is_not_a_file(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/to/directory"
        mock_os_path.exists.return_value = True
        mock_os_path.isdir.return_value = False
        self.explorer.dll.AU3_Send.return_value = 1
        self.explorer.dll.AU3_WinActive.return_value = 1
        
        res = self.explorer.change_directory(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isdir.assert_called_once_with(path)
        mock_check_existing_window.assert_not_called()
        self.explorer.dll.AU3_WinActive.assert_not_called()
        self.explorer.dll.AU3_Send.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, f"path '{path}' is not a directory")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(FileExplorer, "check_existing_window", return_value=None)
    def test_change_directory_failed_to_send_keys_to_open_address_bar(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/to/directory"
        mock_os_path.exists.return_value = True
        mock_os_path.isdir.return_value = True
        self.explorer.dll.AU3_Send.return_value = 0
        self.explorer.dll.AU3_WinActive.return_value = 1
        
        res = self.explorer.change_directory(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isdir.assert_called_once_with(path)
        mock_check_existing_window.assert_called_once()
        self.explorer.dll.AU3_WinActive.assert_not_called()
        self.explorer.dll.AU3_Send.assert_has_calls([
            call("^l", 0)
        ])
        
        self.assertEqual(self.explorer.dll.AU3_Send.call_count, 1)
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not send keys to open address bar")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(FileExplorer, "check_existing_window", return_value=None)
    def test_change_directory_failed_window_inactive_while_filling_address_bar(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/to/directory"
        mock_os_path.exists.return_value = True
        mock_os_path.isdir.return_value = True
        self.explorer.dll.AU3_Send.return_value = 1
        self.explorer.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        
        res = self.explorer.change_directory(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isdir.assert_called_once_with(path)
        mock_check_existing_window.assert_called_once()
        self.explorer.dll.AU3_WinActive.assert_has_calls([call(self.explorer.window_info, "")] * 6)
        self.explorer.dll.AU3_Send.assert_has_calls([
            call("^l", 0),
            *[call(char, 1) for char in "path/"]
        ])
        
        self.assertEqual(self.explorer.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.explorer.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 7)
        self.assertEqual(res, "File Explorer window is inactive")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(FileExplorer, "check_existing_window", return_value=None)
    def test_change_directory_failed_to_send_letter_while_filling_address_bar(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/to/directory"
        mock_os_path.exists.return_value = True
        mock_os_path.isdir.return_value = True
        self.explorer.dll.AU3_Send.side_effect = [1] * 5 + [0]
        self.explorer.dll.AU3_WinActive.return_value = 1
        
        res = self.explorer.change_directory(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isdir.assert_called_once_with(path)
        mock_check_existing_window.assert_called_once()
        self.explorer.dll.AU3_WinActive.assert_has_calls([call(self.explorer.window_info, "")] * 5)
        self.explorer.dll.AU3_Send.assert_has_calls([
            call("^l", 0),
            *[call(char, 1) for char in "path/"]
        ])
        
        self.assertEqual(self.explorer.dll.AU3_WinActive.call_count, 5)
        self.assertEqual(self.explorer.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "could not send / to File Explorer")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(FileExplorer, "check_existing_window", return_value=None)
    def test_change_directory_failed_to_send_key_to_change_directory(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/to/directory"
        mock_os_path.exists.return_value = True
        mock_os_path.isdir.return_value = True
        self.explorer.dll.AU3_Send.side_effect = [1] * len(path) + [1, 0]
        self.explorer.dll.AU3_WinActive.return_value = 1
        
        res = self.explorer.change_directory(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isdir.assert_called_once_with(path)
        mock_check_existing_window.assert_called_once()
        self.explorer.dll.AU3_WinActive.assert_has_calls([call(self.explorer.window_info, "")] * len(path))
        self.explorer.dll.AU3_Send.assert_has_calls([
            call("^l", 0),
            *[call(char, 1) for char in path],
            call("{ENTER}", 0)
        ])
        
        self.assertEqual(self.explorer.dll.AU3_WinActive.call_count, len(path))
        self.assertEqual(self.explorer.dll.AU3_Send.call_count, len(path) + 2)
        self.assertEqual(mock_sleep.call_count, len(path) + 3)
        self.assertEqual(res, "could not send enter key to change directory")

    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(FileExplorer, "change_directory", return_value=None)
    def test_create_directory_success(self, mock_change_directory, mock_os_path, mock_sleep):
        parent_path = "parent/path/example"
        dir_name = "dir/name/example"
        full_path = "parent/path/example/dir/name/example"
        mock_os_path.join.return_value = full_path
        mock_os_path.exists.return_value = False
        self.explorer.dll.AU3_WinActive.return_value = 1
        self.explorer.dll.AU3_Send.return_value = 1

        res = self.explorer.create_directory(parent_path=parent_path, dir_name=dir_name)
        
        mock_os_path.join.assert_called_once_with(parent_path, dir_name)
        mock_os_path.exists.assert_called_once_with(full_path)
        mock_change_directory.assert_called_once_with(parent_path)
        self.explorer.dll.AU3_WinActive.assert_has_calls([call(self.explorer.window_info, "")] * len(dir_name))
        self.explorer.dll.AU3_Send.assert_has_calls([
            call("^+n", 0),
            *[call(char, 1) for char in dir_name],
            call("{ENTER}", 0)
        ])
        
        self.assertEqual(self.explorer.dll.AU3_WinActive.call_count, len(dir_name))
        self.assertEqual(self.explorer.dll.AU3_Send.call_count, len(dir_name) + 2)
        self.assertEqual(mock_sleep.call_count, len(dir_name) + 3)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(FileExplorer, "change_directory", return_value=None)
    def test_create_directory_failed_parent_path_is_not_provided(self, mock_change_directory, mock_os_path, mock_sleep):
        parent_path = ""
        dir_name = "dir/name/example"

        res = self.explorer.create_directory(parent_path=parent_path, dir_name=dir_name)
        
        mock_os_path.join.assert_not_called()
        mock_os_path.exists.assert_not_called()
        mock_change_directory.assert_not_called()
        self.explorer.dll.AU3_WinActive.assert_not_called()
        self.explorer.dll.AU3_Send.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, "parent path must be provided")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(FileExplorer, "change_directory", return_value=None)
    def test_create_directory_failed_directory_name_is_not_provided(self, mock_change_directory, mock_os_path, mock_sleep):
        parent_path = "parent/path/example"
        dir_name = ""

        res = self.explorer.create_directory(parent_path=parent_path, dir_name=dir_name)
        
        mock_os_path.join.assert_not_called()
        mock_os_path.exists.assert_not_called()
        mock_change_directory.assert_not_called()
        self.explorer.dll.AU3_WinActive.assert_not_called()
        self.explorer.dll.AU3_Send.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, "directory name must be provided")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(FileExplorer, "change_directory", return_value=None)
    def test_create_directory_failed_directory_already_exist(self, mock_change_directory, mock_os_path, mock_sleep):
        parent_path = "parent/path/example"
        dir_name = "dir/name/example"
        full_path = "parent/path/example/dir/name/example"
        mock_os_path.join.return_value = full_path
        mock_os_path.exists.return_value = True

        res = self.explorer.create_directory(parent_path=parent_path, dir_name=dir_name)
        
        mock_os_path.join.assert_called_once_with(parent_path, dir_name)
        mock_os_path.exists.assert_called_once_with(full_path)
        mock_change_directory.assert_not_called()
        self.explorer.dll.AU3_WinActive.assert_not_called()
        self.explorer.dll.AU3_Send.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, f"directory '{dir_name}' already exist at '{parent_path}'")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(FileExplorer, "change_directory", return_value=None)
    def test_create_directory_failed_to_send_keys_to_access_new_directory_shortcut(self, mock_change_directory, mock_os_path, mock_sleep):
        parent_path = "parent/path/example"
        dir_name = "dir/name/example"
        full_path = "parent/path/example/dir/name/example"
        mock_os_path.join.return_value = full_path
        mock_os_path.exists.return_value = False
        self.explorer.dll.AU3_WinActive.return_value = 1
        self.explorer.dll.AU3_Send.return_value = 0

        res = self.explorer.create_directory(parent_path=parent_path, dir_name=dir_name)
        
        mock_os_path.join.assert_called_once_with(parent_path, dir_name)
        mock_os_path.exists.assert_called_once_with(full_path)
        mock_change_directory.assert_called_once_with(parent_path)
        self.explorer.dll.AU3_WinActive.assert_not_called()
        self.explorer.dll.AU3_Send.assert_has_calls([
            call("^+n", 0)
        ])
        
        self.assertEqual(self.explorer.dll.AU3_Send.call_count, 1)
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not send keys to access new directory shortcut")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(FileExplorer, "change_directory", return_value=None)
    def test_create_directory_failed_window_inactive_while_filling_directory_name(self, mock_change_directory, mock_os_path, mock_sleep):
        parent_path = "parent/path/example"
        dir_name = "dir/name/example"
        full_path = "parent/path/example/dir/name/example"
        mock_os_path.join.return_value = full_path
        mock_os_path.exists.return_value = False
        self.explorer.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        self.explorer.dll.AU3_Send.return_value = 1

        res = self.explorer.create_directory(parent_path=parent_path, dir_name=dir_name)
        
        mock_os_path.join.assert_called_once_with(parent_path, dir_name)
        mock_os_path.exists.assert_called_once_with(full_path)
        mock_change_directory.assert_called_once_with(parent_path)
        self.explorer.dll.AU3_WinActive.assert_has_calls([call(self.explorer.window_info, "")] * 6)
        self.explorer.dll.AU3_Send.assert_has_calls([
            call("^+n", 0),
            *[call(char, 1) for char in "dir/n"],
        ])
        
        self.assertEqual(self.explorer.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.explorer.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 7)
        self.assertEqual(res,"File Explorer window is inactive")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(FileExplorer, "change_directory", return_value=None)
    def test_create_directory_failed_to_send_letter_while_filling_directory_name(self, mock_change_directory, mock_os_path, mock_sleep):
        parent_path = "parent/path/example"
        dir_name = "dir/name/example"
        full_path = "parent/path/example/dir/name/example"
        mock_os_path.join.return_value = full_path
        mock_os_path.exists.return_value = False
        self.explorer.dll.AU3_WinActive.return_value = 1
        self.explorer.dll.AU3_Send.side_effect = [1] * 5 + [0]

        res = self.explorer.create_directory(parent_path=parent_path, dir_name=dir_name)
        
        mock_os_path.join.assert_called_once_with(parent_path, dir_name)
        mock_os_path.exists.assert_called_once_with(full_path)
        mock_change_directory.assert_called_once_with(parent_path)
        self.explorer.dll.AU3_WinActive.assert_has_calls([call(self.explorer.window_info, "")] * 5)
        self.explorer.dll.AU3_Send.assert_has_calls([
            call("^+n", 0),
            *[call(char, 1) for char in "dir/n"]
        ])
        
        self.assertEqual(self.explorer.dll.AU3_WinActive.call_count, 5)
        self.assertEqual(self.explorer.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res,"could not send n to File Explorer")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(FileExplorer, "change_directory", return_value=None)
    def test_create_directory_failed_to_send_key_to_create_new_directory(self, mock_change_directory, mock_os_path, mock_sleep):
        parent_path = "parent/path/example"
        dir_name = "dir/name/example"
        full_path = "parent/path/example/dir/name/example"
        mock_os_path.join.return_value = full_path
        mock_os_path.exists.return_value = False
        self.explorer.dll.AU3_WinActive.return_value = 1
        self.explorer.dll.AU3_Send.side_effect = [1] * len(dir_name) + [1, 0]

        res = self.explorer.create_directory(parent_path=parent_path, dir_name=dir_name)
        
        mock_os_path.join.assert_called_once_with(parent_path, dir_name)
        mock_os_path.exists.assert_called_once_with(full_path)
        mock_change_directory.assert_called_once_with(parent_path)
        self.explorer.dll.AU3_WinActive.assert_has_calls([call(self.explorer.window_info, "")] * len(dir_name))
        self.explorer.dll.AU3_Send.assert_has_calls([
            call("^+n", 0),
            *[call(char, 1) for char in dir_name],
            call("{ENTER}", 0)
        ])
        
        self.assertEqual(self.explorer.dll.AU3_WinActive.call_count, len(dir_name))
        self.assertEqual(self.explorer.dll.AU3_Send.call_count, len(dir_name) + 2)
        self.assertEqual(mock_sleep.call_count, len(dir_name) + 3)
        self.assertEqual(res, "could not send enter key to create new directory")