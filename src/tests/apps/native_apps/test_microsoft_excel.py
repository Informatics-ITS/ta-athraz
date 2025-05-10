import unittest
from unittest.mock import patch, call, MagicMock
from activities.apps.native_apps.microsoft_excel import MicrosoftExcel

class TestMicrosoftExcel(unittest.TestCase):
    def setUp(self):
        self.excel = MicrosoftExcel()
        self.excel.dll = MagicMock()

    def test_check_existing_window_success(self):
        self.excel.dll.AU3_WinExists.return_value = 1
        self.excel.dll.AU3_WinWaitActive.return_value = 1

        res = self.excel.check_existing_window()

        self.excel.dll.AU3_WinExists.assert_called_once_with(self.excel.window_info, "")
        self.excel.dll.AU3_WinActivate.assert_called_once_with(self.excel.window_info, "")
        self.excel.dll.AU3_WinWaitActive.assert_called_once_with(self.excel.window_info, "", 10)
        
        self.assertIsNone(res)
        
    def test_check_existing_window_failed_window_didnt_exist(self):
        self.excel.dll.AU3_WinExists.return_value = 0
        res = self.excel.check_existing_window()
        
        self.excel.dll.AU3_WinExists.assert_called_once_with(self.excel.window_info, "")
        self.excel.dll.AU3_WinActivate.assert_not_called()
        self.excel.dll.AU3_WinWaitActive.assert_not_called()
        
        self.assertEqual(res, "Microsoft Excel window didn't exist")
        
    def test_check_existing_window_failed_to_wait_window_activation(self):
        self.excel.dll.AU3_WinExists.return_value = 1
        self.excel.dll.AU3_WinActivate.return_value = 1
        self.excel.dll.AU3_WinWaitActive.return_value = 0
        
        res = self.excel.check_existing_window()

        self.excel.dll.AU3_WinExists.assert_called_once_with(self.excel.window_info, "")
        self.excel.dll.AU3_WinActivate.assert_called_once_with(self.excel.window_info, "")
        self.excel.dll.AU3_WinWaitActive.assert_called_once_with(self.excel.window_info, "", 10)
        
        self.assertEqual(res, "could not activate Microsoft Excel window")

    @patch("time.sleep", return_value=None)
    @patch.object(MicrosoftExcel, "_get_executable_path", return_value="/path/to/excel/executable")
    def test_create_window_success(self, mock_get_executable_path, mock_sleep):
        self.excel.dll.AU3_Run.return_value = 1
        self.excel.dll.AU3_WinExists.return_value = 1
        self.excel.dll.AU3_WinGetHandle.return_value = 1234567890
        self.excel.dll.AU3_WinActivate.return_value = 1
        self.excel.dll.AU3_WinWaitActive.return_value = 1
        self.excel.dll.AU3_WinSetState.return_value = 1
        self.excel.dll.AU3_Send.return_value = 1

        res = self.excel.create_window()
        
        mock_get_executable_path.assert_called_once()
        self.excel.dll.AU3_Run.assert_called_once_with("/path/to/excel/executable", "", 1)
        self.excel.dll.AU3_WinExists.assert_called_once()
        self.excel.dll.AU3_WinGetHandle.assert_called_once()
        self.excel.dll.AU3_WinActivate.assert_called_once_with(self.excel.window_info, "")
        self.excel.dll.AU3_WinWaitActive.assert_called_once_with(self.excel.window_info, "", 10)
        self.excel.dll.AU3_WinSetState.assert_called_once_with(self.excel.window_info, "", 3)
        self.excel.dll.AU3_Send.assert_called_once_with("{ENTER}", 0)

        self.assertEqual(mock_sleep.call_count, 3)
        self.assertIsNone(res)

    @patch("time.sleep", return_value=None)
    @patch.object(MicrosoftExcel, "_get_executable_path", return_value=None)
    def test_create_window_failed_to_get_executable_path(self, mock_get_executable_path, mock_sleep):
        res = self.excel.create_window()
        
        mock_get_executable_path.assert_called_once()
        self.excel.dll.AU3_Run.assert_not_called()
        self.excel.dll.AU3_WinExists.assert_not_called()
        self.excel.dll.AU3_WinGetHandle.assert_not_called()
        self.excel.dll.AU3_WinActivate.assert_not_called()
        self.excel.dll.AU3_WinWaitActive.assert_not_called()
        self.excel.dll.AU3_WinSetState.assert_not_called()
        self.excel.dll.AU3_Send.assert_not_called()
        mock_sleep.assert_not_called()
        
        self.assertEqual(res, "could not get Microsoft Excel executable path")

    @patch("time.sleep", return_value=None)
    @patch.object(MicrosoftExcel, "_get_executable_path", return_value="/path/to/excel/executable")
    def test_create_window_failed_to_run(self, mock_get_executable_path, mock_sleep):
        self.excel.dll.AU3_Run.return_value = 0

        res = self.excel.create_window()
        
        mock_get_executable_path.assert_called_once()
        self.excel.dll.AU3_Run.assert_called_once_with("/path/to/excel/executable", "", 1)
        self.excel.dll.AU3_WinExists.assert_not_called()
        self.excel.dll.AU3_WinGetHandle.assert_not_called()
        self.excel.dll.AU3_WinActivate.assert_not_called()
        self.excel.dll.AU3_WinWaitActive.assert_not_called()
        self.excel.dll.AU3_WinSetState.assert_not_called()
        self.excel.dll.AU3_Send.assert_not_called()
        mock_sleep.assert_not_called()
        
        self.assertEqual(res, "could not run Microsoft Excel")

    @patch("time.sleep", return_value=None)
    @patch.object(MicrosoftExcel, "_get_executable_path", return_value="/path/to/excel/executable")
    def test_create_window_failed_window_didnt_exist(self, mock_get_executable_path, mock_sleep):
        self.excel.dll.AU3_Run.return_value = 1
        self.excel.dll.AU3_WinExists.return_value = 0
        res = self.excel.create_window()
        
        mock_get_executable_path.assert_called_once()
        self.excel.dll.AU3_Run.assert_called_once_with("/path/to/excel/executable", "", 1)
        self.excel.dll.AU3_WinExists.assert_called_once()
        self.excel.dll.AU3_WinGetHandle.assert_not_called()
        self.excel.dll.AU3_WinActivate.assert_not_called()
        self.excel.dll.AU3_WinWaitActive.assert_not_called()
        self.excel.dll.AU3_WinSetState.assert_not_called()
        self.excel.dll.AU3_Send.assert_not_called()

        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "Microsoft Excel window didn't exist")

    @patch("time.sleep", return_value=None)
    @patch.object(MicrosoftExcel, "_get_executable_path", return_value="/path/to/excel/executable")
    def test_create_window_failed_to_get_handle(self, mock_get_executable_path, mock_sleep):
        self.excel.dll.AU3_Run.return_value = 1
        self.excel.dll.AU3_WinExists.return_value = 1
        self.excel.dll.AU3_WinGetHandle.return_value = 0
        
        res = self.excel.create_window()
        
        mock_get_executable_path.assert_called_once()
        self.excel.dll.AU3_Run.assert_called_once_with("/path/to/excel/executable", "", 1)
        self.excel.dll.AU3_WinExists.assert_called_once()
        self.excel.dll.AU3_WinGetHandle.assert_called_once()
        self.excel.dll.AU3_WinActivate.assert_not_called()
        self.excel.dll.AU3_WinWaitActive.assert_not_called()
        self.excel.dll.AU3_WinSetState.assert_not_called()
        self.excel.dll.AU3_Send.assert_not_called()


        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not get window handle")

    @patch("time.sleep", return_value=None)
    @patch.object(MicrosoftExcel, "_get_executable_path", return_value="/path/to/excel/executable")
    def test_create_window_failed_to_wait_window_activation(self, mock_get_executable_path, mock_sleep):
        self.excel.dll.AU3_Run.return_value = 1
        self.excel.dll.AU3_WinExists.return_value = 1
        self.excel.dll.AU3_WinGetHandle.return_value = 1234567890
        self.excel.dll.AU3_WinActivate.return_value = 1
        self.excel.dll.AU3_WinWaitActive.return_value = 0

        res = self.excel.create_window()
        
        mock_get_executable_path.assert_called_once()
        self.excel.dll.AU3_Run.assert_called_once_with("/path/to/excel/executable", "", 1)
        self.excel.dll.AU3_WinExists.assert_called_once()
        self.excel.dll.AU3_WinGetHandle.assert_called_once()
        self.excel.dll.AU3_WinActivate.assert_called_once_with(self.excel.window_info, "")
        self.excel.dll.AU3_WinWaitActive.assert_called_once_with(self.excel.window_info, "", 10)
        self.excel.dll.AU3_WinSetState.assert_not_called()
        self.excel.dll.AU3_Send.assert_not_called()

        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not activate Microsoft Excel")

    @patch("time.sleep", return_value=None)
    @patch.object(MicrosoftExcel, "_get_executable_path", return_value="/path/to/excel/executable")
    def test_create_window_failed_to_maximize_window(self, mock_get_executable_path, mock_sleep):
        self.excel.dll.AU3_Run.return_value = 1
        self.excel.dll.AU3_WinExists.return_value = 1
        self.excel.dll.AU3_WinGetHandle.return_value = 1234567890
        self.excel.dll.AU3_WinActivate.return_value = 1
        self.excel.dll.AU3_WinWaitActive.return_value = 1
        self.excel.dll.AU3_WinSetState.return_value = 0

        res = self.excel.create_window()
        
        mock_get_executable_path.assert_called_once()
        self.excel.dll.AU3_Run.assert_called_once_with("/path/to/excel/executable", "", 1)
        self.excel.dll.AU3_WinExists.assert_called_once()
        self.excel.dll.AU3_WinGetHandle.assert_called_once()
        self.excel.dll.AU3_WinActivate.assert_called_once_with(self.excel.window_info, "")
        self.excel.dll.AU3_WinWaitActive.assert_called_once_with(self.excel.window_info, "", 10)
        self.excel.dll.AU3_WinSetState.assert_called_once_with(self.excel.window_info, "", 3)
        self.excel.dll.AU3_Send.assert_not_called()

        self.assertEqual(mock_sleep.call_count, 2)
        self.assertEqual(res, "could not maximize Microsoft Excel window")

    @patch("time.sleep", return_value=None)
    @patch.object(MicrosoftExcel, "_get_executable_path", return_value="/path/to/excel/executable")
    def test_create_window_failed_to_send_key(self, mock_get_executable_path, mock_sleep):
        self.excel.dll.AU3_Run.return_value = 1
        self.excel.dll.AU3_WinExists.return_value = 1
        self.excel.dll.AU3_WinGetHandle.return_value = 1234567890
        self.excel.dll.AU3_WinActivate.return_value = 1
        self.excel.dll.AU3_WinWaitActive.return_value = 1
        self.excel.dll.AU3_WinSetState.return_value = 1
        self.excel.dll.AU3_Send.return_value = 0

        res = self.excel.create_window()
        
        mock_get_executable_path.assert_called_once()
        self.excel.dll.AU3_Run.assert_called_once_with("/path/to/excel/executable", "", 1)
        self.excel.dll.AU3_WinExists.assert_called_once()
        self.excel.dll.AU3_WinGetHandle.assert_called_once()
        self.excel.dll.AU3_WinActivate.assert_called_once_with(self.excel.window_info, "")
        self.excel.dll.AU3_WinWaitActive.assert_called_once_with(self.excel.window_info, "", 10)
        self.excel.dll.AU3_WinSetState.assert_called_once_with(self.excel.window_info, "", 3)
        self.excel.dll.AU3_Send.assert_called_once_with("{ENTER}", 0)

        self.assertEqual(mock_sleep.call_count, 3)
        self.assertEqual(res, "could not send Enter key to create new xlsx")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(MicrosoftExcel, "_get_executable_path", return_value="/path/to/excel/executable")
    def test_open_xlsx_success(self, mock_get_executable_path, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        self.excel.dll.AU3_Run.return_value = 1
        self.excel.dll.AU3_WinExists.return_value = 1
        self.excel.dll.AU3_WinGetHandle.return_value = 1234567890
        self.excel.dll.AU3_WinActivate.return_value = 1
        self.excel.dll.AU3_WinWaitActive.return_value = 1
        self.excel.dll.AU3_WinSetState.return_value = 1
        
        res = self.excel.open_xlsx(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_get_executable_path.assert_called_once()
        self.excel.dll.AU3_Run.assert_called_once_with(f'/path/to/excel/executable "{path}"', "", 1)
        self.excel.dll.AU3_WinExists.assert_called_once()
        self.excel.dll.AU3_WinGetHandle.assert_called_once()
        self.excel.dll.AU3_WinActivate.assert_called_once_with(self.excel.window_info, "")
        self.excel.dll.AU3_WinWaitActive.assert_called_once_with(self.excel.window_info, "", 10)
        self.excel.dll.AU3_WinSetState.assert_called_once_with(self.excel.window_info, "", 3)
        
        self.assertEqual(mock_sleep.call_count, 2)
        self.assertIsNone(res)

    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(MicrosoftExcel, "_get_executable_path")
    def test_open_xlsx_failed_path_is_not_provided(self, mock_get_executable_path, mock_os_path, mock_sleep):
        res = self.excel.open_xlsx(path="")
        
        mock_os_path.exists.assert_not_called()
        mock_os_path.isfile.assert_not_called()
        mock_get_executable_path.assert_not_called()
        self.excel.dll.AU3_Run.assert_not_called()
        self.excel.dll.AU3_WinExists.assert_not_called()
        self.excel.dll.AU3_WinGetHandle.assert_not_called()
        self.excel.dll.AU3_WinActivate.assert_not_called()
        self.excel.dll.AU3_WinWaitActive.assert_not_called()
        self.excel.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 0)
        self.assertEqual(res, "path must be provided")

    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(MicrosoftExcel, "_get_executable_path")  
    def test_open_xlsx_failed_path_didnt_exist(self, mock_get_executable_path, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = False
        
        res = self.excel.open_xlsx(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_not_called()
        mock_get_executable_path.assert_not_called()
        self.excel.dll.AU3_Run.assert_not_called()
        self.excel.dll.AU3_WinExists.assert_not_called()
        self.excel.dll.AU3_WinGetHandle.assert_not_called()
        self.excel.dll.AU3_WinActivate.assert_not_called()
        self.excel.dll.AU3_WinWaitActive.assert_not_called()
        self.excel.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 0)
        self.assertEqual(res, f"file path '{path}' does not exist")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(MicrosoftExcel, "_get_executable_path")
    def test_open_xlsx_failed_path_is_not_a_file(self, mock_get_executable_path, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = False
        
        res = self.excel.open_xlsx(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_get_executable_path.assert_not_called()
        self.excel.dll.AU3_Run.assert_not_called()
        self.excel.dll.AU3_WinExists.assert_not_called()
        self.excel.dll.AU3_WinGetHandle.assert_not_called()
        self.excel.dll.AU3_WinActivate.assert_not_called()
        self.excel.dll.AU3_WinWaitActive.assert_not_called()
        self.excel.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 0)
        self.assertEqual(res, f"path '{path}' is not a file")
               
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(MicrosoftExcel, "_get_executable_path", return_value=None)
    def test_open_xlsx_failed_to_get_executable_path(self, mock_get_executable_path, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        
        res = self.excel.open_xlsx(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_get_executable_path.asset_called_once()
        self.excel.dll.AU3_Run.assert_not_called()
        self.excel.dll.AU3_WinExists.assert_not_called()
        self.excel.dll.AU3_WinGetHandle.assert_not_called()
        self.excel.dll.AU3_WinActivate.assert_not_called()
        self.excel.dll.AU3_WinWaitActive.assert_not_called()
        self.excel.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 0)
        self.assertEqual(res, "could not get Microsoft Excel executable path")
          
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(MicrosoftExcel, "_get_executable_path", return_value="/path/to/excel/executable")
    def test_open_xlsx_failed_to_run(self, mock_get_executable_path, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        self.excel.dll.AU3_Run.return_value = 0
        
        res = self.excel.open_xlsx(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_get_executable_path.assert_called_once()
        self.excel.dll.AU3_Run.assert_called_once_with(f'/path/to/excel/executable "{path}"', "", 1)
        self.excel.dll.AU3_WinExists.assert_not_called()
        self.excel.dll.AU3_WinGetHandle.assert_not_called()
        self.excel.dll.AU3_WinActivate.assert_not_called()
        self.excel.dll.AU3_WinWaitActive.assert_not_called()
        self.excel.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 0)
        self.assertEqual(res, "could not open xlsx file")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(MicrosoftExcel, "_get_executable_path", return_value="/path/to/excel/executable")
    def test_open_xlsx_failed_window_didnt_exist(self, mock_get_executable_path, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        self.excel.dll.AU3_Run.return_value = 1
        self.excel.dll.AU3_WinExists.return_value = 0
        
        res = self.excel.open_xlsx(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_get_executable_path.assert_called_once()
        self.excel.dll.AU3_Run.assert_called_once_with(f'/path/to/excel/executable "{path}"', "", 1)
        self.excel.dll.AU3_WinExists.assert_called_once()
        self.excel.dll.AU3_WinGetHandle.assert_not_called()
        self.excel.dll.AU3_WinActivate.assert_not_called()
        self.excel.dll.AU3_WinWaitActive.assert_not_called()
        self.excel.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "Microsoft Excel window didn't exist")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(MicrosoftExcel, "_get_executable_path", return_value="/path/to/excel/executable")
    def test_open_xlsx_failed_to_get_handle(self, mock_get_executable_path, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        self.excel.dll.AU3_Run.return_value = 1
        self.excel.dll.AU3_WinExists.return_value = 1
        self.excel.dll.AU3_WinGetHandle.return_value = 0
        
        res = self.excel.open_xlsx(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_get_executable_path.assert_called_once()
        self.excel.dll.AU3_Run.assert_called_once_with(f'/path/to/excel/executable "{path}"', "", 1)
        self.excel.dll.AU3_WinExists.assert_called_once()
        self.excel.dll.AU3_WinGetHandle.assert_called_once()
        self.excel.dll.AU3_WinActivate.assert_not_called()
        self.excel.dll.AU3_WinWaitActive.assert_not_called()
        self.excel.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not get window handle")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(MicrosoftExcel, "_get_executable_path", return_value="/path/to/excel/executable")
    def test_open_xlsx_failed_to_wait_window_activation(self, mock_get_executable_path, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        self.excel.dll.AU3_Run.return_value = 1
        self.excel.dll.AU3_WinExists.return_value = 1
        self.excel.dll.AU3_WinGetHandle.return_value = 1234567890
        self.excel.dll.AU3_WinActivate.return_value = 1
        self.excel.dll.AU3_WinWaitActive.return_value = 0
        
        res = self.excel.open_xlsx(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_get_executable_path.assert_called_once()
        self.excel.dll.AU3_Run.assert_called_once_with(f'/path/to/excel/executable "{path}"', "", 1)
        self.excel.dll.AU3_WinExists.assert_called_once()
        self.excel.dll.AU3_WinGetHandle.assert_called_once()
        self.excel.dll.AU3_WinActivate.assert_called_once_with(self.excel.window_info, "")
        self.excel.dll.AU3_WinWaitActive.assert_called_once_with(self.excel.window_info, "", 10)
        self.excel.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not activate Microsoft Excel")
        
    @patch("time.sleep", return_value=None)
    @patch("os.path")
    @patch.object(MicrosoftExcel, "_get_executable_path", return_value="/path/to/excel/executable")
    def test_open_xlsx_failed_to_maximize_window(self, mock_get_executable_path, mock_os_path, mock_sleep):
        path = "path/to/file"
        mock_os_path.exists.return_value = True
        mock_os_path.isfile.return_value = True
        self.excel.dll.AU3_Run.return_value = 1
        self.excel.dll.AU3_WinExists.return_value = 1
        self.excel.dll.AU3_WinGetHandle.return_value = 1234567890
        self.excel.dll.AU3_WinActivate.return_value = 1
        self.excel.dll.AU3_WinWaitActive.return_value = 1
        self.excel.dll.AU3_WinSetState.return_value = 0
        
        res = self.excel.open_xlsx(path=path)
        
        mock_os_path.exists.assert_called_once_with(path)
        mock_os_path.isfile.assert_called_once_with(path)
        mock_get_executable_path.assert_called_once()
        self.excel.dll.AU3_Run.assert_called_once_with(f'/path/to/excel/executable "{path}"', "", 1)
        self.excel.dll.AU3_WinExists.assert_called_once()
        self.excel.dll.AU3_WinGetHandle.assert_called_once()
        self.excel.dll.AU3_WinActivate.assert_called_once_with(self.excel.window_info, "")
        self.excel.dll.AU3_WinWaitActive.assert_called_once_with(self.excel.window_info, "", 10)
        self.excel.dll.AU3_WinSetState.assert_called_once_with(self.excel.window_info, "", 3)
        
        self.assertEqual(mock_sleep.call_count, 2)
        self.assertEqual(res, "could not maximize Microsoft Excel window")
        
    @patch.object(MicrosoftExcel, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_change_cell_success(self, mock_sleep, mock_check_existing_window):
        target_cell = "F5"
        self.excel.dll.AU3_Send.return_value = 1
        self.excel.dll.AU3_WinExists.return_value = 1
        self.excel.dll.AU3_WinActivate.return_value = 1
        self.excel.dll.AU3_WinWaitActive.return_value = 1
        
        res = self.excel.change_cell(target_cell=target_cell)
        
        mock_check_existing_window.assert_called_once()
        self.excel.dll.AU3_Send.assert_has_calls([
            call("^g", 0),
            call(target_cell, 1),
            call("{ENTER}", 0),
        ])
        self.excel.dll.AU3_WinExists.assert_called_once_with(self.excel.goto_window_info, "")
        self.excel.dll.AU3_WinActivate.assert_called_once_with(self.excel.goto_window_info, "")
        self.excel.dll.AU3_WinWaitActive.assert_called_once_with(self.excel.goto_window_info, "", 10)
        
        self.assertEqual(self.excel.dll.AU3_Send.call_count, 3)
        self.assertEqual(mock_sleep.call_count, 4)
        self.assertIsNone(res)

    @patch.object(MicrosoftExcel, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_change_chell_failed_to_send_keys_to_activate_goto_modal(self, mock_sleep, mock_check_existing_window):
        target_cell = "F5"
        self.excel.dll.AU3_Send.return_value = 0
        
        res = self.excel.change_cell(target_cell=target_cell)
        
        mock_check_existing_window.assert_called_once()
        self.excel.dll.AU3_Send.assert_has_calls([
            call("^g", 0)
        ])
        self.excel.dll.AU3_WinExists.assert_not_called()
        self.excel.dll.AU3_WinActivate.assert_not_called()
        self.excel.dll.AU3_WinWaitActive.assert_not_called()
        
        self.assertEqual(self.excel.dll.AU3_Send.call_count, 1)
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not send keys to open Go To window")
        
    @patch.object(MicrosoftExcel, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_change_cell_failed_goto_modal_window_didnt_exist(self, mock_sleep, mock_check_existing_window):
        target_cell = "F5"
        self.excel.dll.AU3_Send.return_value = 1
        self.excel.dll.AU3_WinExists.return_value = 0
        
        res = self.excel.change_cell(target_cell=target_cell)
        
        mock_check_existing_window.assert_called_once()
        self.excel.dll.AU3_Send.assert_has_calls([
            call("^g", 0)
        ])
        self.excel.dll.AU3_WinExists.assert_called_once_with(self.excel.goto_window_info, "")
        self.excel.dll.AU3_WinActivate.assert_not_called()
        self.excel.dll.AU3_WinWaitActive.assert_not_called()
        
        self.assertEqual(self.excel.dll.AU3_Send.call_count, 1)
        self.assertEqual(mock_sleep.call_count, 2)
        self.assertEqual(res, "Go To window didn't exist")

    @patch.object(MicrosoftExcel, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_change_cell_failed_to_wait_window_activation(self, mock_sleep, mock_check_existing_window):
        target_cell = "F5"
        self.excel.dll.AU3_Send.return_value = 1
        self.excel.dll.AU3_WinExists.return_value = 1
        self.excel.dll.AU3_WinActivate.return_value = 1
        self.excel.dll.AU3_WinWaitActive.return_value = 0
        
        res = self.excel.change_cell(target_cell=target_cell)
        
        mock_check_existing_window.assert_called_once()
        self.excel.dll.AU3_Send.assert_has_calls([
            call("^g", 0)
        ])
        self.excel.dll.AU3_WinExists.assert_called_once_with(self.excel.goto_window_info, "")
        self.excel.dll.AU3_WinActivate.assert_called_once_with(self.excel.goto_window_info, "")
        self.excel.dll.AU3_WinWaitActive.assert_called_once_with(self.excel.goto_window_info, "", 10)
        
        self.assertEqual(self.excel.dll.AU3_Send.call_count, 1)
        self.assertEqual(mock_sleep.call_count, 2)
        self.assertEqual(res, "could not activate Go To window")
        
    @patch.object(MicrosoftExcel, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_change_cell_failed_to_send_target_cell(self, mock_sleep, mock_check_existing_window):
        target_cell = "F5"
        self.excel.dll.AU3_Send.side_effect = [1, 0]
        self.excel.dll.AU3_WinExists.return_value = 1
        self.excel.dll.AU3_WinActivate.return_value = 1
        self.excel.dll.AU3_WinWaitActive.return_value = 1
        
        res = self.excel.change_cell(target_cell=target_cell)
        
        mock_check_existing_window.assert_called_once()
        self.excel.dll.AU3_Send.assert_has_calls([
            call("^g", 0),
            call(target_cell, 1)
        ])
        self.excel.dll.AU3_WinExists.assert_called_once_with(self.excel.goto_window_info, "")
        self.excel.dll.AU3_WinActivate.assert_called_once_with(self.excel.goto_window_info, "")
        self.excel.dll.AU3_WinWaitActive.assert_called_once_with(self.excel.goto_window_info, "", 10)
        
        self.assertEqual(self.excel.dll.AU3_Send.call_count, 2)
        self.assertEqual(mock_sleep.call_count, 3)
        self.assertEqual(res, "could not send target cell to Go To window")
        
    @patch.object(MicrosoftExcel, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_change_cell_failed_to_send_key_to_change_cell(self, mock_sleep, mock_check_existing_window):
        target_cell = "F5"
        self.excel.dll.AU3_Send.side_effect = [1, 1, 0]
        self.excel.dll.AU3_WinExists.return_value = 1
        self.excel.dll.AU3_WinActivate.return_value = 1
        self.excel.dll.AU3_WinWaitActive.return_value = 1
        
        res = self.excel.change_cell(target_cell=target_cell)
        
        mock_check_existing_window.assert_called_once()
        self.excel.dll.AU3_Send.assert_has_calls([
            call("^g", 0),
            call(target_cell, 1),
            call("{ENTER}", 0),
        ])
        self.excel.dll.AU3_WinExists.assert_called_once_with(self.excel.goto_window_info, "")
        self.excel.dll.AU3_WinActivate.assert_called_once_with(self.excel.goto_window_info, "")
        self.excel.dll.AU3_WinWaitActive.assert_called_once_with(self.excel.goto_window_info, "", 10)
        
        self.assertEqual(self.excel.dll.AU3_Send.call_count, 3)
        self.assertEqual(mock_sleep.call_count, 4)
        self.assertEqual(res, "could not send Enter key to change cell")
        
    @patch.object(MicrosoftExcel, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_create_sheet_success(self, mock_sleep, mock_check_existing_window):
        self.excel.dll.AU3_Send.return_value = 1
        
        res = self.excel.create_sheet()
        
        mock_check_existing_window.assert_called_once()
        self.excel.dll.AU3_Send.assert_called_once_with("+{F11}", 0)
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertIsNone(res)
        
    @patch.object(MicrosoftExcel, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_create_sheet_failed_to_send_keys(self, mock_sleep, mock_check_existing_window):
        self.excel.dll.AU3_Send.return_value = 0
        
        res = self.excel.create_sheet()
        
        mock_check_existing_window.assert_called_once()
        self.excel.dll.AU3_Send.assert_called_once_with("+{F11}", 0)
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not send keys to create new sheet")

    @patch.object(MicrosoftExcel, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_rename_sheet_success(self, mock_sleep, mock_check_existing_window):
        new_name = "sheet updated"
        self.excel.dll.AU3_Send.return_value = 1
        self.excel.dll.AU3_WinActive.return_value = 1
        
        res = self.excel.rename_sheet(new_name=new_name)
        
        mock_check_existing_window.assert_called_once()
        self.excel.dll.AU3_Send.assert_has_calls([
            call("!hor", 0),
            *[call(char, 1) for char in new_name],
            call("{ENTER}", 0),
        ])
        self.excel.dll.AU3_WinActive.assert_has_calls([call(self.excel.window_info, "")] * len(new_name))
        
        self.assertEqual(self.excel.dll.AU3_WinActive.call_count, len(new_name))
        self.assertEqual(self.excel.dll.AU3_Send.call_count, len(new_name) + 2)
        self.assertEqual(mock_sleep.call_count, len(new_name) + 3)
        self.assertIsNone(res)

    @patch.object(MicrosoftExcel, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_rename_sheet_failed_new_name_is_not_provided(self, mock_sleep, mock_check_existing_window):
        new_name = ""
        
        res = self.excel.rename_sheet(new_name=new_name)
        
        mock_check_existing_window.assert_not_called()
        self.excel.dll.AU3_Send.assert_not_called()
        self.excel.dll.AU3_WinActive.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, "new name must be provided")

    @patch.object(MicrosoftExcel, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_rename_sheet_failed_to_send_keys_to_rename_sheet(self, mock_sleep, mock_check_existing_window):
        new_name = "sheet updated"
        self.excel.dll.AU3_Send.return_value = 0
        
        res = self.excel.rename_sheet(new_name=new_name)
        
        mock_check_existing_window.assert_called_once()
        self.excel.dll.AU3_Send.assert_called_once_with("!hor", 0)
        self.excel.dll.AU3_WinActive.assert_not_called()

        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not send keys to rename sheet")

    @patch.object(MicrosoftExcel, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_rename_sheet_failed_window_inactive_while_filling_new_name(self, mock_sleep, mock_check_existing_window):
        new_name = "sheet updated"
        self.excel.dll.AU3_Send.return_value = 1
        self.excel.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        
        res = self.excel.rename_sheet(new_name=new_name)
        
        mock_check_existing_window.assert_called_once()
        self.excel.dll.AU3_Send.assert_has_calls([
            call("!hor", 0),
            *[call(char, 1) for char in "sheet"],
        ])
        self.excel.dll.AU3_WinActive.assert_has_calls([call(self.excel.window_info, "")] * 6)
        
        self.assertEqual(self.excel.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.excel.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 7)
        self.assertEqual(res, "Microsoft Excel window is inactive")

    @patch.object(MicrosoftExcel, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_rename_sheet_failed_to_send_letter_while_filling_new_name(self, mock_sleep, mock_check_existing_window):
        new_name = "sheet updated"
        self.excel.dll.AU3_Send.side_effect = [1] * 5 + [0]
        self.excel.dll.AU3_WinActive.return_value = 1
        
        res = self.excel.rename_sheet(new_name=new_name)
        
        mock_check_existing_window.assert_called_once()
        self.excel.dll.AU3_Send.assert_has_calls([
            call("!hor", 0),
            *[call(char, 1) for char in "sheet"],
        ])
        self.excel.dll.AU3_WinActive.assert_has_calls([call(self.excel.window_info, "")] * 5)
        
        self.assertEqual(self.excel.dll.AU3_WinActive.call_count, 5)
        self.assertEqual(self.excel.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "could not send t to Microsoft Excel")

    @patch.object(MicrosoftExcel, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_rename_sheet_failed_to_send_key_to_save_new_name(self, mock_sleep, mock_check_existing_window):
        new_name = "sheet updated"
        self.excel.dll.AU3_Send.side_effect = [1] * (len(new_name) + 1) + [0]
        self.excel.dll.AU3_WinActive.return_value = 1
        
        res = self.excel.rename_sheet(new_name=new_name)
        
        mock_check_existing_window.assert_called_once()
        self.excel.dll.AU3_Send.assert_has_calls([
            call("!hor", 0),
            *[call(char, 1) for char in new_name],
            call("{ENTER}", 0),
        ])
        self.excel.dll.AU3_WinActive.assert_has_calls([call(self.excel.window_info, "")] * len(new_name))
        
        self.assertEqual(self.excel.dll.AU3_WinActive.call_count, len(new_name))
        self.assertEqual(self.excel.dll.AU3_Send.call_count, len(new_name) + 2)
        self.assertEqual(mock_sleep.call_count, len(new_name) + 3)
        self.assertEqual(res, "could not send Enter key to save new name")

    @patch.object(MicrosoftExcel, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_change_sheet_success(self, mock_sleep, mock_check_existing_window):
        direction = "left"
        count = 10
        self.excel.dll.AU3_WinActive.return_value = 1
        self.excel.dll.AU3_Send.return_value = 1
        
        res = self.excel.change_sheet(direction=direction, count=count)
        
        mock_check_existing_window.assert_called_once()
        self.excel.dll.AU3_Send.assert_has_calls([call("^{PGUP}", 0)] * count)
        self.excel.dll.AU3_WinActive.assert_has_calls([call(self.excel.window_info, "")] * count)
        
        self.assertEqual(self.excel.dll.AU3_WinActive.call_count, count)
        self.assertEqual(self.excel.dll.AU3_Send.call_count, count)
        self.assertEqual(mock_sleep.call_count, count + 1)
        self.assertIsNone(res)

    @patch.object(MicrosoftExcel, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_change_sheet_failed_invalid_change_sheet_direction(self, mock_sleep, mock_check_existing_window):
        direction = "abc"
        count = 10
        self.excel.dll.AU3_WinActive.return_value = 1
        self.excel.dll.AU3_Send.return_value = 1
        
        res = self.excel.change_sheet(direction=direction, count=count)
        
        mock_check_existing_window.assert_not_called()
        self.excel.dll.AU3_Send.assert_not_called()
        self.excel.dll.AU3_WinActive.assert_not_called()
        mock_sleep.assert_not_called()
        self.assertEqual(res, "invalid change sheet direction")

    @patch.object(MicrosoftExcel, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_change_sheet_failed_window_inactive_while_changing_sheets(self, mock_sleep, mock_check_existing_window):
        direction = "left"
        count = 10
        self.excel.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        self.excel.dll.AU3_Send.return_value = 1
        
        res = self.excel.change_sheet(direction=direction, count=count)
        
        mock_check_existing_window.assert_called_once()
        self.excel.dll.AU3_Send.assert_has_calls([call("^{PGUP}", 0)] * 5)
        self.excel.dll.AU3_WinActive.assert_has_calls([call(self.excel.window_info, "")] * 6)
        
        self.assertEqual(self.excel.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.excel.dll.AU3_Send.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "Microsoft Excel window is inactive")

    @patch.object(MicrosoftExcel, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_change_sheet_failed_to_send_keys_to_change_sheets_left(self, mock_sleep, mock_check_existing_window):
        direction = "left"
        count = 10
        self.excel.dll.AU3_WinActive.return_value = 1
        self.excel.dll.AU3_Send.side_effect = [1] * 5 + [0]
        
        res = self.excel.change_sheet(direction=direction, count=count)
        
        mock_check_existing_window.assert_called_once()
        self.excel.dll.AU3_Send.assert_has_calls([call("^{PGUP}", 0)] * 6)
        self.excel.dll.AU3_WinActive.assert_has_calls([call(self.excel.window_info, "")] * 6)
        
        self.assertEqual(self.excel.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.excel.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "could not send keys to change sheets")

    @patch.object(MicrosoftExcel, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_change_sheet_failed_to_send_keys_to_change_sheets_right(self, mock_sleep, mock_check_existing_window):
        direction = "right"
        count = 10
        self.excel.dll.AU3_WinActive.return_value = 1
        self.excel.dll.AU3_Send.side_effect = [1] * 5 + [0]
        
        res = self.excel.change_sheet(direction=direction, count=count)
        
        mock_check_existing_window.assert_called_once()
        self.excel.dll.AU3_Send.assert_has_calls([call("^{PGDN}", 0)] * 6)
        self.excel.dll.AU3_WinActive.assert_has_calls([call(self.excel.window_info, "")] * 6)
        
        self.assertEqual(self.excel.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.excel.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "could not send keys to change sheets")

    @patch.object(MicrosoftExcel, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_write_cell_success(self, mock_sleep, mock_check_existing_window):
        text = "lorem ipsum dolor sit amet"
        self.excel.dll.AU3_WinActive.return_value = 1
        self.excel.dll.AU3_Send.return_value = 1
        
        res = self.excel.write_cell(text=text)
        
        mock_check_existing_window.assert_called_once()
        self.excel.dll.AU3_WinActive.assert_has_calls([call(self.excel.window_info, "")] * len(text))
        self.excel.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in text],
            call("{ENTER}", 0),
        ])
        
        self.assertEqual(self.excel.dll.AU3_WinActive.call_count, len(text))
        self.assertEqual(self.excel.dll.AU3_Send.call_count, len(text) + 1)
        self.assertEqual(mock_sleep.call_count, len(text) + 2)
        self.assertIsNone(res)

    @patch.object(MicrosoftExcel, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_write_cell_failed_window_inactive_while_writing_cell(self, mock_sleep, mock_check_existing_window):
        text = "lorem ipsum dolor sit amet"
        self.excel.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        self.excel.dll.AU3_Send.return_value = 1
        
        res = self.excel.write_cell(text=text)
        
        mock_check_existing_window.assert_called_once()
        self.excel.dll.AU3_WinActive.assert_has_calls([call(self.excel.window_info, "")] * 6)
        self.excel.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in "lorem"]
        ])
        
        self.assertEqual(self.excel.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.excel.dll.AU3_Send.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "Microsoft Excel window is inactive")

    @patch.object(MicrosoftExcel, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_write_cell_failed_to_send_letter_while_writing_cell(self, mock_sleep, mock_check_existing_window):
        text = "lorem ipsum dolor sit amet"
        self.excel.dll.AU3_WinActive.return_value = 1
        self.excel.dll.AU3_Send.side_effect = [1] * 4 + [0]
        
        res = self.excel.write_cell(text=text)
        
        mock_check_existing_window.assert_called_once()
        self.excel.dll.AU3_WinActive.assert_has_calls([call(self.excel.window_info, "")] * 5)
        self.excel.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in "lorem"],
        ])
        
        self.assertEqual(self.excel.dll.AU3_WinActive.call_count, 5)
        self.assertEqual(self.excel.dll.AU3_Send.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 5)
        self.assertEqual(res, "could not send m to Microsoft Excel")

    @patch.object(MicrosoftExcel, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_write_cell_failed_to_send_key_to_confirm_cell(self, mock_sleep, mock_check_existing_window):
        text = "lorem ipsum dolor sit amet"
        self.excel.dll.AU3_WinActive.return_value = 1
        self.excel.dll.AU3_Send.side_effect = [1] * len(text) + [0]
        
        res = self.excel.write_cell(text=text)
        
        mock_check_existing_window.assert_called_once()
        self.excel.dll.AU3_WinActive.assert_has_calls([call(self.excel.window_info, "")] * len(text))
        self.excel.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in text],
            call("{ENTER}", 0),
        ])
        
        self.assertEqual(self.excel.dll.AU3_WinActive.call_count, len(text))
        self.assertEqual(self.excel.dll.AU3_Send.call_count, len(text) + 1)
        self.assertEqual(mock_sleep.call_count, len(text) + 2)
        self.assertEqual(res, "could not send enter key to write cell")

    @patch.object(MicrosoftExcel, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_scroll_success(self, mock_sleep, mock_check_existing_window):
        direction = "down"
        clicks = 30
        scroll_delay = 0.07
        self.excel.dll.AU3_WinActive.return_value = 1
        self.excel.dll.AU3_MouseWheel.return_value = 1
        
        res = self.excel.scroll(direction=direction, clicks=clicks, scroll_delay=scroll_delay)
        
        mock_check_existing_window.assert_called_once()
        self.excel.dll.AU3_WinActive.assert_has_calls([call(self.excel.window_info, "")] * clicks)
        self.excel.dll.AU3_MouseWheel.assert_has_calls([call(direction, 1)] * clicks)
        
        self.assertEqual(self.excel.dll.AU3_WinActive.call_count, clicks)
        self.assertEqual(self.excel.dll.AU3_MouseWheel.call_count, clicks)
        self.assertEqual(mock_sleep.call_count, clicks + 1)
        self.assertIsNone(res)

    @patch.object(MicrosoftExcel, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_scroll_failed_invalid_scroll_direction(self, mock_sleep, mock_check_existing_window):
        direction = "abc"
        clicks = 30
        scroll_delay = 0.07
        
        res = self.excel.scroll(direction=direction, clicks=clicks, scroll_delay=scroll_delay)
        
        mock_check_existing_window.assert_not_called()
        self.excel.dll.AU3_WinActive.assert_not_called()
        self.excel.dll.AU3_MouseWheel.assert_not_called()
        mock_sleep.assert_not_called()

        self.assertEqual(res, "invalid scroll direction")

    @patch.object(MicrosoftExcel, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_scroll_failed_window_inactive_while_scrolling(self, mock_sleep, mock_check_existing_window):
        direction = "down"
        clicks = 30
        scroll_delay = 0.07
        self.excel.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        self.excel.dll.AU3_MouseWheel.return_value = 1
        
        res = self.excel.scroll(direction=direction, clicks=clicks, scroll_delay=scroll_delay)
        
        mock_check_existing_window.assert_called_once()
        self.excel.dll.AU3_WinActive.assert_has_calls([call(self.excel.window_info, "")] * 6)
        self.excel.dll.AU3_MouseWheel.assert_has_calls([call(direction, 1)] * 5)
        
        self.assertEqual(self.excel.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.excel.dll.AU3_MouseWheel.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 5)
        self.assertEqual(res, "Microsoft Word window is inactive")

    @patch.object(MicrosoftExcel, "check_existing_window", return_value=None) 
    @patch("time.sleep", return_value=None)
    def test_scroll_failed_window_inactive_while_scrolling(self, mock_sleep, mock_check_existing_window):
        direction = "down"
        clicks = 30
        scroll_delay = 0.07
        self.excel.dll.AU3_WinActive.return_value = 1
        self.excel.dll.AU3_MouseWheel.side_effect = [1] * 5 + [0]
        
        res = self.excel.scroll(direction=direction, clicks=clicks, scroll_delay=scroll_delay)
        
        mock_check_existing_window.assert_called_once()
        self.excel.dll.AU3_WinActive.assert_has_calls([call(self.excel.window_info, "")] * 6)
        self.excel.dll.AU3_MouseWheel.assert_has_calls([call(direction, 1)] * 6)
        
        self.assertEqual(self.excel.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.excel.dll.AU3_MouseWheel.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 7)
        self.assertEqual(res, "could not scroll mouse wheel")
        
    @patch.object(MicrosoftExcel, "check_existing_window", return_value=None) 
    def test_save_xlsx_success(self, mock_check_existing_window):
        self.excel.dll.AU3_Send.return_value = 1
        
        res = self.excel.save_xlsx()
        
        mock_check_existing_window.assert_called_once() 
        self.excel.dll.AU3_Send.assert_called_once_with("^s", 0)
        self.assertIsNone(res)
        
    @patch.object(MicrosoftExcel, "check_existing_window", return_value=None) 
    def test_save_xlsx_failed(self, mock_check_existing_window):
        self.excel.dll.AU3_Send.return_value = 0
        
        res = self.excel.save_xlsx()
        
        mock_check_existing_window.assert_called_once() 
        self.excel.dll.AU3_Send.assert_called_once_with("^s", 0)
        self.assertEqual(res, "could not send ctrl+s to save xlsx file")