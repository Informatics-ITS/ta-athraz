import unittest
from unittest.mock import patch, call, MagicMock
from activities.files.command_prompt import CommandPrompt

class TestCommandPrompt(unittest.TestCase):
    def setUp(self):
        self.cmd = CommandPrompt()
        self.cmd.dll = MagicMock()
        
    def test_check_existing_window_success(self):
        self.cmd.dll.AU3_WinExists.return_value = 1
        self.cmd.dll.AU3_WinWaitActive.return_value = 1

        res = self.cmd.check_existing_window()

        self.cmd.dll.AU3_WinExists.assert_called_once_with(self.cmd.window_info, "")
        self.cmd.dll.AU3_WinActivate.assert_called_once_with(self.cmd.window_info, "")
        self.cmd.dll.AU3_WinWaitActive.assert_called_once_with(self.cmd.window_info, "", 10)
        
        self.assertIsNone(res)
        
    def test_check_existing_window_failed_window_didnt_exist(self):
        self.cmd.dll.AU3_WinExists.return_value = 0
        res = self.cmd.check_existing_window()
        
        self.cmd.dll.AU3_WinExists.assert_called_once_with(self.cmd.window_info, "")
        self.cmd.dll.AU3_WinActivate.assert_not_called()
        self.cmd.dll.AU3_WinWaitActive.assert_not_called()
        
        self.assertEqual(res, "Command Prompt window didn't exist")
        
    def test_check_existing_window_failed_to_wait_window_activation(self):
        self.cmd.dll.AU3_WinExists.return_value = 1
        self.cmd.dll.AU3_WinActivate.return_value = 1
        self.cmd.dll.AU3_WinWaitActive.return_value = 0
        
        res = self.cmd.check_existing_window()

        self.cmd.dll.AU3_WinExists.assert_called_once_with(self.cmd.window_info, "")
        self.cmd.dll.AU3_WinActivate.assert_called_once_with(self.cmd.window_info, "")
        self.cmd.dll.AU3_WinWaitActive.assert_called_once_with(self.cmd.window_info, "", 10)
        
        self.assertEqual(res, "could not activate Command Prompt window")
        
    @patch("time.sleep", return_value=None)
    def test_create_window_success(self, mock_sleep):
        self.cmd.dll.AU3_Run.return_value = 1
        self.cmd.dll.AU3_WinExists.return_value = 1
        self.cmd.dll.AU3_WinGetHandle.return_value = 1234567890
        self.cmd.dll.AU3_WinActivate.return_value = 1
        self.cmd.dll.AU3_WinWaitActive.return_value = 1
        self.cmd.dll.AU3_WinSetState.return_value = 1

        res = self.cmd.create_window()

        self.cmd.dll.AU3_Run.assert_called_once_with('cmd.exe /c start "Command Prompt"', "", 1)
        self.cmd.dll.AU3_WinExists.assert_called_once()
        self.cmd.dll.AU3_WinGetHandle.assert_called_once()
        self.cmd.dll.AU3_WinActivate.assert_called_once_with(self.cmd.window_info, "")
        self.cmd.dll.AU3_WinWaitActive.assert_called_once_with(self.cmd.window_info, "", 10)
        self.cmd.dll.AU3_WinSetState.assert_called_once_with(self.cmd.window_info, "", 3)

        self.assertEqual(mock_sleep.call_count, 2)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    def test_create_window_failed_to_run(self, mock_sleep):
        self.cmd.dll.AU3_Run.return_value = 0
        res = self.cmd.create_window()

        self.cmd.dll.AU3_Run.assert_called_once_with('cmd.exe /c start "Command Prompt"', "", 1)
        self.cmd.dll.AU3_WinExists.assert_not_called()
        self.cmd.dll.AU3_WinGetHandle.assert_not_called()
        self.cmd.dll.AU3_WinActivate.assert_not_called()
        self.cmd.dll.AU3_WinWaitActive.assert_not_called()
        self.cmd.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(res, "could not run Command Prompt")
        
    @patch("time.sleep", return_value=None)
    def test_create_window_failed_window_didnt_exist(self, mock_sleep):
        self.cmd.dll.AU3_Run.return_value = 1
        self.cmd.dll.AU3_WinExists.return_value = 0

        res = self.cmd.create_window()

        self.cmd.dll.AU3_Run.assert_called_once_with('cmd.exe /c start "Command Prompt"', "", 1)
        self.cmd.dll.AU3_WinExists.assert_called_once()
        self.cmd.dll.AU3_WinGetHandle.assert_not_called()
        self.cmd.dll.AU3_WinActivate.assert_not_called()
        self.cmd.dll.AU3_WinWaitActive.assert_not_called()
        self.cmd.dll.AU3_WinSetState.assert_not_called()

        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "Command Prompt window didn't exist")
        
    @patch("time.sleep", return_value=None)
    def test_create_window_failed_to_get_handle(self, mock_sleep):
        self.cmd.dll.AU3_Run.return_value = 1
        self.cmd.dll.AU3_WinExists.return_value = 1
        self.cmd.dll.AU3_WinGetHandle.return_value = 0

        res = self.cmd.create_window()

        self.cmd.dll.AU3_Run.assert_called_once_with('cmd.exe /c start "Command Prompt"', "", 1)
        self.cmd.dll.AU3_WinExists.assert_called_once()
        self.cmd.dll.AU3_WinGetHandle.assert_called_once()
        self.cmd.dll.AU3_WinActivate.assert_not_called()
        self.cmd.dll.AU3_WinWaitActive.assert_not_called()
        self.cmd.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not get window handle")
        
    @patch("time.sleep", return_value=None)
    def test_create_window_failed_to_wait_window_activation(self, mock_sleep):
        self.cmd.dll.AU3_Run.return_value = 1
        self.cmd.dll.AU3_WinExists.return_value = 1
        self.cmd.dll.AU3_WinGetHandle.return_value = 1234567890
        self.cmd.dll.AU3_WinActivate.return_value = 1
        self.cmd.dll.AU3_WinWaitActive.return_value = 0

        res = self.cmd.create_window()

        self.cmd.dll.AU3_Run.assert_called_once_with('cmd.exe /c start "Command Prompt"', "", 1)
        self.cmd.dll.AU3_WinExists.assert_called_once()
        self.cmd.dll.AU3_WinGetHandle.assert_called_once()
        self.cmd.dll.AU3_WinActivate.assert_called_once_with(self.cmd.window_info, "")
        self.cmd.dll.AU3_WinWaitActive.assert_called_once_with(self.cmd.window_info, "", 10)
        self.cmd.dll.AU3_WinSetState.assert_not_called()

        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not activate Command Prompt")
        
    @patch("time.sleep", return_value=None)
    def test_create_window_failed_to_maximize_window(self, mock_sleep):
        self.cmd.dll.AU3_Run.return_value = 1
        self.cmd.dll.AU3_WinExists.return_value = 1
        self.cmd.dll.AU3_WinGetHandle.return_value = 1234567890
        self.cmd.dll.AU3_WinActivate.return_value = 1
        self.cmd.dll.AU3_WinWaitActive.return_value = 1
        self.cmd.dll.AU3_WinSetState.return_value = 0

        res = self.cmd.create_window()

        self.cmd.dll.AU3_Run.assert_called_once_with('cmd.exe /c start "Command Prompt"', "", 1)
        self.cmd.dll.AU3_WinExists.assert_called_once()
        self.cmd.dll.AU3_WinGetHandle.assert_called_once()
        self.cmd.dll.AU3_WinActivate.assert_called_once_with(self.cmd.window_info, "")
        self.cmd.dll.AU3_WinWaitActive.assert_called_once_with(self.cmd.window_info, "", 10)
        self.cmd.dll.AU3_WinSetState.assert_called_once_with(self.cmd.window_info, "", 3)

        self.assertEqual(mock_sleep.call_count, 2)
        self.assertEqual(res, "could not maximize Command Prompt window")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_open_file_success(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/to/file"
        cmdline = f'start "" "{path}"'
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        self.cmd.dll.AU3_WinActive.return_value = 1
        self.cmd.dll.AU3_Send.return_value = 1
        
        res = self.cmd.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_check_existing_window.assert_called_once()
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * len(cmdline))
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in cmdline],
            call("{ENTER}", 0)
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, len(cmdline))
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, len(cmdline) + 1)
        self.assertEqual(mock_sleep.call_count, len(cmdline) + 2)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_open_file_failed_path_is_not_provided(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = ""
        
        res = self.cmd.open_file(path=path)
        
        mock_os_path.exists.assert_not_called()
        mock_os_path.isfile.assert_not_called()
        mock_check_existing_window.assert_not_called()
        self.cmd.dll.AU3_WinActive.assert_not_called()
        self.cmd.dll.AU3_Send.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, "file path must be provided")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_open_file_failed_path_didnt_exist(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = False
        
        res = self.cmd.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_not_called()
        mock_check_existing_window.assert_not_called()
        self.cmd.dll.AU3_WinActive.assert_not_called()
        self.cmd.dll.AU3_Send.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, f"file path '{path}' does not exist")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_open_file_failed_path_is_not_a_file(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = False

        res = self.cmd.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_check_existing_window.assert_not_called()
        self.cmd.dll.AU3_WinActive.assert_not_called()
        self.cmd.dll.AU3_Send.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, f"path '{path}' is not a file")

    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_open_file_failed_window_inactive_while_sending_command_line(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        self.cmd.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        self.cmd.dll.AU3_Send.return_value = 1
        
        res = self.cmd.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_check_existing_window.assert_called_once()
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * 6)
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in "start"],
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "Command Prompt window is inactive")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_open_file_failed_to_send_command_line_letter(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        self.cmd.dll.AU3_WinActive.return_value = 1
        self.cmd.dll.AU3_Send.side_effect = [1] * 4 + [0]
        
        res = self.cmd.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_check_existing_window.assert_called_once()
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * 5)
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in "start"],
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, 5)
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 5)
        self.assertEqual(res, "could not send t to Command Prompt")

    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_open_file_failed_to_send_key_to_execute_command(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/to/file"
        cmdline = f'start "" "{path}"'
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        self.cmd.dll.AU3_WinActive.return_value = 1
        self.cmd.dll.AU3_Send.side_effect = [1] * len(cmdline) + [0]
        
        res = self.cmd.open_file(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_check_existing_window.assert_called_once()
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * len(cmdline))
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in cmdline],
            call("{ENTER}", 0)
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, len(cmdline))
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, len(cmdline) + 1)
        self.assertEqual(mock_sleep.call_count, len(cmdline) + 2)
        self.assertEqual(res, "could not send enter key to execute command")

    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_change_directory_success(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/to/directory"
        cmdline = f'cd "{path}"'
        mock_os_path.exists.return_value = True
        mock_os_path.isdir.return_value = True
        self.cmd.dll.AU3_WinActive.return_value = 1
        self.cmd.dll.AU3_Send.return_value = 1
        
        res = self.cmd.change_directory(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isdir.assert_called_once_with(path)
        mock_check_existing_window.assert_called_once()
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * len(cmdline))
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in cmdline],
            call("{ENTER}", 0)
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, len(cmdline))
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, len(cmdline) + 1)
        self.assertEqual(mock_sleep.call_count, len(cmdline) + 2)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_change_directory_failed_path_is_not_provided(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = ""
        
        res = self.cmd.change_directory(path=path)
        
        mock_os_path.exists.assert_not_called()
        mock_os_path.isdir.assert_not_called()
        mock_check_existing_window.assert_not_called()
        self.cmd.dll.AU3_WinActive.assert_not_called()
        self.cmd.dll.AU3_Send.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, "directory path must be provided")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_change_directory_failed_path_didnt_exist(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/to/directory"
        mock_os_path.exists.return_value = False
        
        res = self.cmd.change_directory(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isdir.assert_not_called()
        mock_check_existing_window.assert_not_called()
        self.cmd.dll.AU3_WinActive.assert_not_called()
        self.cmd.dll.AU3_Send.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, f"directory '{path}' does not exist")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_change_directory_failed_path_is_not_a_directory(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/to/directory"
        mock_os_path.exists.return_value = True
        mock_os_path.isdir.return_value = False

        res = self.cmd.change_directory(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isdir.assert_called_once_with(path)
        mock_check_existing_window.assert_not_called()
        self.cmd.dll.AU3_WinActive.assert_not_called()
        self.cmd.dll.AU3_Send.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, f"path '{path}' is not a directory")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_change_directory_failed_window_inactive_while_sending_command_line(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isdir.return_value = True
        self.cmd.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        self.cmd.dll.AU3_Send.return_value = 1
        
        res = self.cmd.change_directory(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isdir.assert_called_once_with(path)
        mock_check_existing_window.assert_called_once()
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * 6)
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in 'cd "p'],
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "Command Prompt window is inactive")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_change_directory_failed_to_send_command_line_letter(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/to/directory"
        mock_os_path.exists.return_value = True
        mock_os_path.isdir.return_value = True
        self.cmd.dll.AU3_WinActive.return_value = 1
        self.cmd.dll.AU3_Send.side_effect = [1] * 4 + [0]
        
        res = self.cmd.change_directory(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isdir.assert_called_once_with(path)
        mock_check_existing_window.assert_called_once()
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * 5)
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in 'cd "p'],
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, 5)
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 5)
        self.assertEqual(res, "could not send p to Command Prompt")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_change_directory_failed_to_send_key_to_execute_command(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/to/directory"
        cmdline = f'cd "{path}"'
        mock_os_path.exists.return_value = True
        mock_os_path.isdir.return_value = True
        self.cmd.dll.AU3_WinActive.return_value = 1
        self.cmd.dll.AU3_Send.side_effect = [1] * len(cmdline) + [0]
        
        res = self.cmd.change_directory(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isdir.assert_called_once_with(path)
        mock_check_existing_window.assert_called_once()
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * len(cmdline))
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in cmdline],
            call("{ENTER}", 0)
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, len(cmdline))
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, len(cmdline) + 1)
        self.assertEqual(mock_sleep.call_count, len(cmdline) + 2)
        self.assertEqual(res, "could not send enter key to execute command")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "change_directory", return_value=None)
    def test_create_directory_success(self, mock_change_directory, mock_os_path, mock_sleep):
        parent_path = "parent/path/example"
        dir_name = "dir/name/example"
        full_path = "parent/path/example/dir/name/example"
        cmdline = f'mkdir "{dir_name}"'
        mock_os_path.join.return_value = full_path
        mock_os_path.exists.return_value = False
        self.cmd.dll.AU3_WinActive.return_value = 1
        self.cmd.dll.AU3_Send.return_value = 1

        res = self.cmd.create_directory(parent_path=parent_path, dir_name=dir_name)
        
        mock_os_path.join.assert_called_once_with(parent_path, dir_name)
        mock_os_path.exists.assert_called_once_with(full_path)
        mock_change_directory.assert_called_once_with(parent_path)
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * len(cmdline))
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in cmdline],
            call("{ENTER}", 0)
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, len(cmdline))
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, len(cmdline) + 1)
        self.assertEqual(mock_sleep.call_count, len(cmdline) + 2)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "change_directory", return_value=None)
    def test_create_directory_failed_parent_path_is_not_provided(self, mock_change_directory, mock_os_path, mock_sleep):
        parent_path = ""
        dir_name = "dir/name/example"

        res = self.cmd.create_directory(parent_path=parent_path, dir_name=dir_name)
        
        mock_os_path.join.assert_not_called()
        mock_os_path.exists.assert_not_called()
        mock_change_directory.assert_not_called()
        self.cmd.dll.AU3_WinActive.assert_not_called()
        self.cmd.dll.AU3_Send.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, "parent path must be provided")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "change_directory", return_value=None)
    def test_create_directory_failed_directory_name_is_not_provided(self, mock_change_directory, mock_os_path, mock_sleep):
        parent_path = "parent/path/example"
        dir_name = ""

        res = self.cmd.create_directory(parent_path=parent_path, dir_name=dir_name)
        
        mock_os_path.join.assert_not_called()
        mock_os_path.exists.assert_not_called()
        mock_change_directory.assert_not_called()
        self.cmd.dll.AU3_WinActive.assert_not_called()
        self.cmd.dll.AU3_Send.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, "directory name must be provided")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "change_directory", return_value=None)
    def test_create_directory_failed_directory_already_exist(self, mock_change_directory, mock_os_path, mock_sleep):
        parent_path = "parent/path/example"
        dir_name = "dir/name/example"
        full_path = "parent/path/example/dir/name/example"
        mock_os_path.join.return_value = full_path
        mock_os_path.exists.return_value = True

        res = self.cmd.create_directory(parent_path=parent_path, dir_name=dir_name)
        
        mock_os_path.join.assert_called_once_with(parent_path, dir_name)
        mock_os_path.exists.assert_called_once_with(full_path)
        mock_change_directory.assert_not_called()
        self.cmd.dll.AU3_WinActive.assert_not_called()
        self.cmd.dll.AU3_Send.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, f"directory '{dir_name}' already exist at '{parent_path}'")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "change_directory", return_value=None)
    def test_create_directory_failed_window_inactive_while_sending_command_line(self, mock_change_directory, mock_os_path, mock_sleep):
        parent_path = "parent/path/example"
        dir_name = "dir/name/example"
        full_path = "parent/path/example/dir/name/example"
        mock_os_path.join.return_value = full_path
        mock_os_path.exists.return_value = False
        self.cmd.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        self.cmd.dll.AU3_Send.return_value = 1

        res = self.cmd.create_directory(parent_path=parent_path, dir_name=dir_name)
        
        mock_os_path.join.assert_called_once_with(parent_path, dir_name)
        mock_os_path.exists.assert_called_once_with(full_path)
        mock_change_directory.assert_called_once_with(parent_path)
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * 6)
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in "mkdir"],
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "Command Prompt window is inactive")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "change_directory", return_value=None)
    def test_create_directory_failed_to_send_command_line_letter(self, mock_change_directory, mock_os_path, mock_sleep):
        parent_path = "parent/path/example"
        dir_name = "dir/name/example"
        full_path = "parent/path/example/dir/name/example"
        mock_os_path.join.return_value = full_path
        mock_os_path.exists.return_value = False
        self.cmd.dll.AU3_WinActive.return_value = 1
        self.cmd.dll.AU3_Send.side_effect = [1] * 4 + [0]

        res = self.cmd.create_directory(parent_path=parent_path, dir_name=dir_name)
        
        mock_os_path.join.assert_called_once_with(parent_path, dir_name)
        mock_os_path.exists.assert_called_once_with(full_path)
        mock_change_directory.assert_called_once_with(parent_path)
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * 5)
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in "mkdir"],
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, 5)
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 5)
        self.assertEqual(res, "could not send r to Command Prompt")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "change_directory", return_value=None)
    def test_create_directory_failed_to_send_key_to_execute_command(self, mock_change_directory, mock_os_path, mock_sleep):
        parent_path = "parent/path/example"
        dir_name = "dir/name/example"
        full_path = "parent/path/example/dir/name/example"
        cmdline = f'mkdir "{dir_name}"'
        mock_os_path.join.return_value = full_path
        mock_os_path.exists.return_value = False
        self.cmd.dll.AU3_WinActive.return_value = 1
        self.cmd.dll.AU3_Send.side_effect = [1] * len(cmdline) + [0]

        res = self.cmd.create_directory(parent_path=parent_path, dir_name=dir_name)
        
        mock_os_path.join.assert_called_once_with(parent_path, dir_name)
        mock_os_path.exists.assert_called_once_with(full_path)
        mock_change_directory.assert_called_once_with(parent_path)
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * len(cmdline))
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in cmdline],
            call("{ENTER}", 0)
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, len(cmdline))
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, len(cmdline) + 1)
        self.assertEqual(mock_sleep.call_count, len(cmdline) + 2)
        self.assertEqual(res, "could not send enter key to execute command")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_copy_success_source_path_directory(self, mock_check_existing_window, mock_os_path, mock_sleep):
        source_path = "source/path/example"
        destination_path = "destination/path/example"
        cmdline = f'xcopy "{source_path}" "{destination_path}" /e /i /h'
        mock_os_path.exists.return_value = True
        mock_os_path.isdir.return_value = True
        self.cmd.dll.AU3_WinActive.return_value = 1
        self.cmd.dll.AU3_Send.return_value = 1

        res = self.cmd.copy(source_path=source_path, destination_path=destination_path)
        
        mock_os_path.exists.assert_called_once_with(source_path)
        mock_check_existing_window.assert_called_once()
        mock_os_path.isdir.assert_called_once_with(source_path)
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * len(cmdline))
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in cmdline],
            call("{ENTER}", 0)
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, len(cmdline))
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, len(cmdline) + 1)
        self.assertEqual(mock_sleep.call_count, len(cmdline) + 2)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_copy_success_source_path_file(self, mock_check_existing_window, mock_os_path, mock_sleep):
        source_path = "source/path/example"
        destination_path = "destination/path/example"
        cmdline = f'copy "{source_path}" "{destination_path}"'
        mock_os_path.exists.return_value = True
        mock_os_path.isdir.return_value = False
        self.cmd.dll.AU3_WinActive.return_value = 1
        self.cmd.dll.AU3_Send.return_value = 1

        res = self.cmd.copy(source_path=source_path, destination_path=destination_path)
        
        mock_os_path.exists.assert_called_once_with(source_path)
        mock_check_existing_window.assert_called_once()
        mock_os_path.isdir.assert_called_once_with(source_path)
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * len(cmdline))
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in cmdline],
            call("{ENTER}", 0)
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, len(cmdline))
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, len(cmdline) + 1)
        self.assertEqual(mock_sleep.call_count, len(cmdline) + 2)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_copy_failed_source_path_is_not_provided(self, mock_check_existing_window, mock_os_path, mock_sleep):
        source_path = ""
        destination_path = "destination/path/example"

        res = self.cmd.copy(source_path=source_path, destination_path=destination_path)
        
        mock_os_path.exists.assert_not_called()
        mock_check_existing_window.assert_not_called()
        mock_os_path.isdir.assert_not_called()
        self.cmd.dll.AU3_WinActive.assert_not_called()
        self.cmd.dll.AU3_Send.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, "source path must be provided")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_copy_failed_destination_path_is_not_provided(self, mock_check_existing_window, mock_os_path, mock_sleep):
        source_path = "source/path/example"
        destination_path = ""
        cmdline = f'xcopy "{source_path}" "{destination_path}" /e /i /h'

        res = self.cmd.copy(source_path=source_path, destination_path=destination_path)
        
        mock_os_path.exists.assert_not_called()
        mock_check_existing_window.assert_not_called()
        mock_os_path.isdir.assert_not_called()
        self.cmd.dll.AU3_WinActive.assert_not_called()
        self.cmd.dll.AU3_Send.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, "destination path must be provided")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_copy_failed_source_path_doesnt_exist(self, mock_check_existing_window, mock_os_path, mock_sleep):
        source_path = "source/path/example"
        destination_path = "destination/path/example"
        mock_os_path.exists.return_value = False

        res = self.cmd.copy(source_path=source_path, destination_path=destination_path)
        
        mock_os_path.exists.assert_called_once_with(source_path)
        mock_check_existing_window.assert_not_called()
        mock_os_path.isdir.assert_not_called()
        self.cmd.dll.AU3_WinActive.assert_not_called()
        self.cmd.dll.AU3_Send.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, f"source path '{source_path}' does not exist")

    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_copy_failed_window_inactive_while_sending_command_line(self, mock_check_existing_window, mock_os_path, mock_sleep):
        source_path = "source/path/example"
        destination_path = "destination/path/example"
        mock_os_path.exists.return_value = True
        mock_os_path.isdir.return_value = True
        self.cmd.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        self.cmd.dll.AU3_Send.return_value = 1

        res = self.cmd.copy(source_path=source_path, destination_path=destination_path)
        
        mock_os_path.exists.assert_called_once_with(source_path)
        mock_check_existing_window.assert_called_once()
        mock_os_path.isdir.assert_called_once_with(source_path)
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * 6)
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in "xcopy"],
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "Command Prompt window is inactive")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_copy_failed_to_send_command_line_letter(self, mock_check_existing_window, mock_os_path, mock_sleep):
        source_path = "source/path/example"
        destination_path = "destination/path/example"
        mock_os_path.exists.return_value = True
        mock_os_path.isdir.return_value = True
        self.cmd.dll.AU3_WinActive.return_value = 1
        self.cmd.dll.AU3_Send.side_effect = [1] * 4 + [0]

        res = self.cmd.copy(source_path=source_path, destination_path=destination_path)
        
        mock_os_path.exists.assert_called_once_with(source_path)
        mock_check_existing_window.assert_called_once()
        mock_os_path.isdir.assert_called_once_with(source_path)
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * 5)
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in "xcopy"],
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, 5)
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 5)
        self.assertEqual(res, "could not send y to Command Prompt")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_copy_failed_to_send_key_to_execute_command(self, mock_check_existing_window, mock_os_path, mock_sleep):
        source_path = "source/path/example"
        destination_path = "destination/path/example"
        cmdline = f'xcopy "{source_path}" "{destination_path}" /e /i /h'
        mock_os_path.exists.return_value = True
        mock_os_path.isdir.return_value = True
        self.cmd.dll.AU3_WinActive.return_value = 1
        self.cmd.dll.AU3_Send.side_effect = [1] * len(cmdline) + [0]

        res = self.cmd.copy(source_path=source_path, destination_path=destination_path)
        
        mock_os_path.exists.assert_called_once_with(source_path)
        mock_check_existing_window.assert_called_once()
        mock_os_path.isdir.assert_called_once_with(source_path)
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * len(cmdline))
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in cmdline],
            call("{ENTER}", 0)
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, len(cmdline))
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, len(cmdline) + 1)
        self.assertEqual(mock_sleep.call_count, len(cmdline) + 2)
        self.assertEqual(res, "could not send enter key to execute command")

    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_rename_success(self, mock_check_existing_window, mock_os_path, mock_sleep):
        old_name = "old/name/example"
        new_name = "new_name"
        cmdline = f'ren "{old_name}" "{new_name}"'
        mock_os_path.exists.return_value = True
        self.cmd.dll.AU3_WinActive.return_value = 1
        self.cmd.dll.AU3_Send.return_value = 1

        res = self.cmd.rename(old_name=old_name, new_name=new_name)
        
        mock_os_path.exists.assert_called_once_with(old_name)
        mock_check_existing_window.assert_called_once()
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * len(cmdline))
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in cmdline],
            call("{ENTER}", 0)
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, len(cmdline))
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, len(cmdline) + 1)
        self.assertEqual(mock_sleep.call_count, len(cmdline) + 2)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_rename_failed_old_name_is_not_provided(self, mock_check_existing_window, mock_os_path, mock_sleep):
        old_name = ""
        new_name = "new_name"

        res = self.cmd.rename(old_name=old_name, new_name=new_name)
        
        mock_os_path.exists.assert_not_called()
        mock_check_existing_window.assert_not_called()
        self.cmd.dll.AU3_WinActive.assert_not_called()
        self.cmd.dll.AU3_Send.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, "old name must be provided")

    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_rename_failed_new_name_is_not_provided(self, mock_check_existing_window, mock_os_path, mock_sleep):
        old_name = "old/name/example"
        new_name = ""

        res = self.cmd.rename(old_name=old_name, new_name=new_name)
        
        mock_os_path.exists.assert_not_called()
        mock_check_existing_window.assert_not_called()
        self.cmd.dll.AU3_WinActive.assert_not_called()
        self.cmd.dll.AU3_Send.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, "new name must be provided")

    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_rename_failed_old_name_doesnt_exist(self, mock_check_existing_window, mock_os_path, mock_sleep):
        old_name = "old/name/example"
        new_name = "new_name"
        mock_os_path.exists.return_value = False

        res = self.cmd.rename(old_name=old_name, new_name=new_name)
        
        mock_os_path.exists.assert_called_once_with(old_name)
        mock_check_existing_window.assert_not_called()
        self.cmd.dll.AU3_WinActive.assert_not_called()
        self.cmd.dll.AU3_Send.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, f"'{old_name}' does not exist")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_rename_failed_window_inactive_while_sending_command_line(self, mock_check_existing_window, mock_os_path, mock_sleep):
        old_name = "old/name/example"
        new_name = "new_name"
        mock_os_path.exists.return_value = True
        self.cmd.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        self.cmd.dll.AU3_Send.return_value = 1

        res = self.cmd.rename(old_name=old_name, new_name=new_name)
        
        mock_os_path.exists.assert_called_once_with(old_name)
        mock_check_existing_window.assert_called_once()
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * 6)
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in 'ren "']
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "Command Prompt window is inactive")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_rename_failed_to_send_command_line_letter(self, mock_check_existing_window, mock_os_path, mock_sleep):
        old_name = "old/name/example"
        new_name = "new_name"
        mock_os_path.exists.return_value = True
        self.cmd.dll.AU3_WinActive.return_value = 1
        self.cmd.dll.AU3_Send.side_effect = [1] * 5 + [0]

        res = self.cmd.rename(old_name=old_name, new_name=new_name)
        
        mock_os_path.exists.assert_called_once_with(old_name)
        mock_check_existing_window.assert_called_once()
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * 6)
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in 'ren "o'],
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "could not send o to Command Prompt")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_rename_failed_to_send_key_to_execute_command(self, mock_check_existing_window, mock_os_path, mock_sleep):
        old_name = "old/name/example"
        new_name = "new_name"
        cmdline = f'ren "{old_name}" "{new_name}"'
        mock_os_path.exists.return_value = True
        self.cmd.dll.AU3_WinActive.return_value = 1
        self.cmd.dll.AU3_Send.side_effect = [1] * len(cmdline) + [0]

        res = self.cmd.rename(old_name=old_name, new_name=new_name)
        
        mock_os_path.exists.assert_called_once_with(old_name)
        mock_check_existing_window.assert_called_once()
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * len(cmdline))
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in cmdline],
            call("{ENTER}", 0)
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, len(cmdline))
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, len(cmdline) + 1)
        self.assertEqual(mock_sleep.call_count, len(cmdline) + 2)
        self.assertEqual(res, "could not send enter key to execute command")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_delete_success_path_directory(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/example"
        cmdline = f'rmdir /s /q "{path}"'
        mock_os_path.exists.return_value = True
        mock_os_path.isdir.return_value = True
        self.cmd.dll.AU3_WinActive.return_value = 1
        self.cmd.dll.AU3_Send.return_value = 1

        res = self.cmd.delete(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_check_existing_window.assert_called_once()
        mock_os_path.isdir.assert_called_once_with(path)
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * len(cmdline))
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in cmdline],
            call("{ENTER}", 0)
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, len(cmdline))
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, len(cmdline) + 1)
        self.assertEqual(mock_sleep.call_count, len(cmdline) + 2)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_delete_success_path_file(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/example"
        cmdline = f'del "{path}"'
        mock_os_path.exists.return_value = True
        mock_os_path.isdir.return_value = False
        self.cmd.dll.AU3_WinActive.return_value = 1
        self.cmd.dll.AU3_Send.return_value = 1

        res = self.cmd.delete(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_check_existing_window.assert_called_once()
        mock_os_path.isdir.assert_called_once_with(path)
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * len(cmdline))
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in cmdline],
            call("{ENTER}", 0)
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, len(cmdline))
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, len(cmdline) + 1)
        self.assertEqual(mock_sleep.call_count, len(cmdline) + 2)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_delete_failed_path_is_not_provided(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = ""

        res = self.cmd.delete(path=path)
        
        mock_os_path.exists.assert_not_called()
        mock_check_existing_window.assert_not_called()
        mock_os_path.isdir.assert_not_called()
        self.cmd.dll.AU3_WinActive.assert_not_called()
        self.cmd.dll.AU3_Send.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, "path must be provided")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_delete_failed_path_doesnt_exist(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/example"
        mock_os_path.exists.return_value = False

        res = self.cmd.delete(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_check_existing_window.assert_not_called()
        mock_os_path.isdir.assert_not_called()
        self.cmd.dll.AU3_WinActive.assert_not_called()
        self.cmd.dll.AU3_Send.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, f"path '{path}' does not exist")
           
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_delete_failed_window_inactive_while_sending_command_line(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/example"
        mock_os_path.exists.return_value = True
        mock_os_path.isdir.return_value = True
        self.cmd.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        self.cmd.dll.AU3_Send.return_value = 1

        res = self.cmd.delete(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_check_existing_window.assert_called_once()
        mock_os_path.isdir.assert_called_once_with(path)
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * 6)
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in "rmdir"]
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "Command Prompt window is inactive")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_delete_failed_to_send_command_line_letter(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/example"
        mock_os_path.exists.return_value = True
        mock_os_path.isdir.return_value = True
        self.cmd.dll.AU3_WinActive.return_value = 1
        self.cmd.dll.AU3_Send.side_effect = [1] * 4 + [0]

        res = self.cmd.delete(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_check_existing_window.assert_called_once()
        mock_os_path.isdir.assert_called_once_with(path)
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * 5)
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in "rmdir"]
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, 5)
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 5)
        self.assertEqual(res, "could not send r to Command Prompt")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_delete_failed_to_send_key_to_execute_command(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/example"
        cmdline = f'rmdir /s /q "{path}"'
        mock_os_path.exists.return_value = True
        mock_os_path.isdir.return_value = True
        self.cmd.dll.AU3_WinActive.return_value = 1
        self.cmd.dll.AU3_Send.side_effect = [1] * len(cmdline) + [0]

        res = self.cmd.delete(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_check_existing_window.assert_called_once()
        mock_os_path.isdir.assert_called_once_with(path)
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * len(cmdline))
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in cmdline],
            call("{ENTER}", 0)
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, len(cmdline))
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, len(cmdline) + 1)
        self.assertEqual(mock_sleep.call_count, len(cmdline) + 2)
        self.assertEqual(res, "could not send enter key to execute command")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_move_success(self, mock_check_existing_window, mock_os_path, mock_sleep):
        source_path = "source/path/example"
        destination_path = "destination/path/example"
        cmdline = f'move "{source_path}" "{destination_path}"'
        mock_os_path.exists.return_value = True
        self.cmd.dll.AU3_WinActive.return_value = 1
        self.cmd.dll.AU3_Send.return_value = 1

        res = self.cmd.move(source_path=source_path, destination_path=destination_path)
        
        mock_os_path.exists.assert_called_once_with(source_path)
        mock_check_existing_window.assert_called_once()
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * len(cmdline))
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in cmdline],
            call("{ENTER}", 0)
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, len(cmdline))
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, len(cmdline) + 1)
        self.assertEqual(mock_sleep.call_count, len(cmdline) + 2)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_move_failed_source_path_is_not_provided(self, mock_check_existing_window, mock_os_path, mock_sleep):
        source_path = ""
        destination_path = "destination/path/example"
        cmdline = f'move "{source_path}" "{destination_path}"'

        res = self.cmd.move(source_path=source_path, destination_path=destination_path)
        
        mock_os_path.exists.assert_not_called()
        mock_check_existing_window.assert_not_called()
        self.cmd.dll.AU3_WinActive.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, "source path must be provided")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_move_failed_destination_path_is_not_provided(self, mock_check_existing_window, mock_os_path, mock_sleep):
        source_path = "source/path/example"
        destination_path = ""

        res = self.cmd.move(source_path=source_path, destination_path=destination_path)
        
        mock_os_path.exists.assert_not_called()
        mock_check_existing_window.assert_not_called()
        self.cmd.dll.AU3_WinActive.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, "destination path must be provided")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_move_failed_source_path_doesnt_exist(self, mock_check_existing_window, mock_os_path, mock_sleep):
        source_path = "source/path/example"
        destination_path = "destination/path/example"
        mock_os_path.exists.return_value = False

        res = self.cmd.move(source_path=source_path, destination_path=destination_path)
        
        mock_os_path.exists.assert_called_once_with(source_path)
        mock_check_existing_window.assert_not_called()
        self.cmd.dll.AU3_WinActive.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, f"source path '{source_path}' does not exist")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_move_failed_window_inactive_while_sending_command_line(self, mock_check_existing_window, mock_os_path, mock_sleep):
        source_path = "source/path/example"
        destination_path = "destination/path/example"
        mock_os_path.exists.return_value = True
        self.cmd.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        self.cmd.dll.AU3_Send.return_value = 1

        res = self.cmd.move(source_path=source_path, destination_path=destination_path)
        
        mock_os_path.exists.assert_called_once_with(source_path)
        mock_check_existing_window.assert_called_once()
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * 6)
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in "move "]
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "Command Prompt window is inactive")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_move_failed_to_send_command_line_letter(self, mock_check_existing_window, mock_os_path, mock_sleep):
        source_path = "source/path/example"
        destination_path = "destination/path/example"
        mock_os_path.exists.return_value = True
        self.cmd.dll.AU3_WinActive.return_value = 1
        self.cmd.dll.AU3_Send.side_effect = [1] * 5 + [0]

        res = self.cmd.move(source_path=source_path, destination_path=destination_path)
        
        mock_os_path.exists.assert_called_once_with(source_path)
        mock_check_existing_window.assert_called_once()
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * 6)
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in 'move "']
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, 'could not send " to Command Prompt')
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_move_failed_to_send_key_to_execute_command(self, mock_check_existing_window, mock_os_path, mock_sleep):
        source_path = "source/path/example"
        destination_path = "destination/path/example"
        cmdline = f'move "{source_path}" "{destination_path}"'
        mock_os_path.exists.return_value = True
        self.cmd.dll.AU3_WinActive.return_value = 1
        self.cmd.dll.AU3_Send.side_effect = [1] * len(cmdline) + [0]

        res = self.cmd.move(source_path=source_path, destination_path=destination_path)
        
        mock_os_path.exists.assert_called_once_with(source_path)
        mock_check_existing_window.assert_called_once()
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * len(cmdline))
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in cmdline],
            call("{ENTER}", 0)
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, len(cmdline))
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, len(cmdline) + 1)
        self.assertEqual(mock_sleep.call_count, len(cmdline) + 2)
        self.assertEqual(res, "could not send enter key to execute command")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_list_success_with_path(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/example"
        cmdline = f'dir "{path}"'
        mock_os_path.exists.return_value = True
        self.cmd.dll.AU3_WinActive.return_value = 1
        self.cmd.dll.AU3_Send.return_value = 1

        res = self.cmd.list(path=path)
        
        mock_check_existing_window.assert_called_once()
        mock_os_path.exists.assert_called_once_with(path)
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * len(cmdline))
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in cmdline],
            call("{ENTER}", 0)
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, len(cmdline))
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, len(cmdline) + 1)
        self.assertEqual(mock_sleep.call_count, len(cmdline) + 2)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_list_success_without_path(self, mock_check_existing_window, mock_os_path, mock_sleep):
        cmdline = "dir"
        self.cmd.dll.AU3_WinActive.return_value = 1
        self.cmd.dll.AU3_Send.return_value = 1

        res = self.cmd.list()
        
        mock_check_existing_window.assert_called_once()
        mock_os_path.exists.assert_not_called()
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * len(cmdline))
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in cmdline],
            call("{ENTER}", 0)
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, len(cmdline))
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, len(cmdline) + 1)
        self.assertEqual(mock_sleep.call_count, len(cmdline) + 2)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_list_failed_path_doesnt_exist(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/example"
        mock_os_path.exists.return_value = False

        res = self.cmd.list(path=path)
        
        mock_check_existing_window.assert_called_once()
        mock_os_path.exists.assert_called_once_with(path)
        self.cmd.dll.AU3_WinActive.assert_not_called()
        self.cmd.dll.AU3_Send.assert_not_called()
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, f"path '{path}' does not exist")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_list_failed_window_inactive_while_sending_command_line(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/example"
        mock_os_path.exists.return_value = True
        self.cmd.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        self.cmd.dll.AU3_Send.return_value = 1

        res = self.cmd.list(path=path)
        
        mock_check_existing_window.assert_called_once()
        mock_os_path.exists.assert_called_once_with(path)
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * 6)
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in 'dir "']
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "Command Prompt window is inactive")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_list_failed_to_send_command_line_letter(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/example"
        mock_os_path.exists.return_value = True
        self.cmd.dll.AU3_WinActive.return_value = 1
        self.cmd.dll.AU3_Send.side_effect = [1] * 5 + [0]

        res = self.cmd.list(path=path)
        
        mock_check_existing_window.assert_called_once()
        mock_os_path.exists.assert_called_once_with(path)
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * 6)
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in 'dir "p']
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "could not send p to Command Prompt")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(CommandPrompt, "check_existing_window", return_value=None)
    def test_list_failed_to_send_key_execute_command(self, mock_check_existing_window, mock_os_path, mock_sleep):
        path = "path/example"
        cmdline = f'dir "{path}"'
        mock_os_path.exists.return_value = True
        self.cmd.dll.AU3_WinActive.return_value = 1
        self.cmd.dll.AU3_Send.side_effect = [1] * len(cmdline) + [0]

        res = self.cmd.list(path=path)
        
        mock_check_existing_window.assert_called_once()
        mock_os_path.exists.assert_called_once_with(path)
        self.cmd.dll.AU3_WinActive.assert_has_calls([call(self.cmd.window_info, "")] * len(cmdline))
        self.cmd.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in cmdline],
            call("{ENTER}", 0)
        ])
        
        self.assertEqual(self.cmd.dll.AU3_WinActive.call_count, len(cmdline))
        self.assertEqual(self.cmd.dll.AU3_Send.call_count, len(cmdline) + 1)
        self.assertEqual(mock_sleep.call_count, len(cmdline) + 2)
        self.assertEqual(res, "could not send enter key to execute command")