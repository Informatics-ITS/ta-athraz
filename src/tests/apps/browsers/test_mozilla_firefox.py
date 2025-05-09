import unittest
from unittest.mock import patch, call, MagicMock
from activities.apps.browsers.mozilla_firefox import MozillaFirefox

class TestMozillaFirefox(unittest.TestCase):
    def setUp(self):
        self.firefox = MozillaFirefox()
        self.firefox.dll = MagicMock()
        
    def test_check_existing_window_success(self):
        self.firefox.dll.AU3_WinExists.return_value = 1
        self.firefox.dll.AU3_WinWaitActive.return_value = 1

        res = self.firefox.check_existing_window()

        self.firefox.dll.AU3_WinExists.assert_called_once_with(self.firefox.window_info, "")
        self.firefox.dll.AU3_WinActivate.assert_called_once_with(self.firefox.window_info, "")
        self.firefox.dll.AU3_WinWaitActive.assert_called_once_with(self.firefox.window_info, "", 10)
        
        self.assertIsNone(res)
        
    def test_check_existing_window_failed_window_didnt_exist(self):
        self.firefox.dll.AU3_WinExists.return_value = 0
        res = self.firefox.check_existing_window()
        
        self.firefox.dll.AU3_WinExists.assert_called_once_with(self.firefox.window_info, "")
        self.firefox.dll.AU3_WinActivate.assert_not_called()
        self.firefox.dll.AU3_WinWaitActive.assert_not_called()
        
        self.assertEqual(res, "Mozilla Firefox window didn't exist")
        
    def test_check_existing_window_failed_to_wait_window_activation(self):
        self.firefox.dll.AU3_WinExists.return_value = 1
        self.firefox.dll.AU3_WinActivate.return_value = 1
        self.firefox.dll.AU3_WinWaitActive.return_value = 0
        
        res = self.firefox.check_existing_window()

        self.firefox.dll.AU3_WinExists.assert_called_once_with(self.firefox.window_info, "")
        self.firefox.dll.AU3_WinActivate.assert_called_once_with(self.firefox.window_info, "")
        self.firefox.dll.AU3_WinWaitActive.assert_called_once_with(self.firefox.window_info, "", 10)
        
        self.assertEqual(res, "could not activate Mozilla Firefox window")

    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "_get_executable_path", return_value="/path/to/firefox/executable")
    def test_create_window_success(self, mock_get_executable_path, mock_sleep):
        self.firefox.dll.AU3_Run.return_value = 1
        self.firefox.dll.AU3_WinExists.return_value = 1
        self.firefox.dll.AU3_WinGetHandle.return_value = 1234567890
        self.firefox.dll.AU3_WinActivate.return_value = 1
        self.firefox.dll.AU3_WinWaitActive.return_value = 1
        self.firefox.dll.AU3_WinSetState.return_value = 1

        res = self.firefox.create_window()
        
        mock_get_executable_path.assert_called_once()
        self.firefox.dll.AU3_Run.assert_called_once_with("/path/to/firefox/executable", "", 1)
        self.firefox.dll.AU3_WinExists.assert_called_once()
        self.firefox.dll.AU3_WinGetHandle.assert_called_once()
        self.firefox.dll.AU3_WinActivate.assert_called_once_with(self.firefox.window_info, "")
        self.firefox.dll.AU3_WinWaitActive.assert_called_once_with(self.firefox.window_info, "", 10)
        self.firefox.dll.AU3_WinSetState.assert_called_once_with(self.firefox.window_info, "", 3)

        self.assertEqual(mock_sleep.call_count, 2)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "_get_executable_path", return_value=None)
    def test_create_window_failed_to_get_executable_path(self, mock_get_executable_path, mock_sleep):
        res = self.firefox.create_window()
        
        mock_get_executable_path.assert_called_once()
        self.firefox.dll.AU3_Run.assert_not_called()
        self.firefox.dll.AU3_WinExists.assert_not_called()
        self.firefox.dll.AU3_WinGetHandle.assert_not_called()
        self.firefox.dll.AU3_WinActivate.assert_not_called()
        self.firefox.dll.AU3_WinWaitActive.assert_not_called()
        self.firefox.dll.AU3_WinSetState.assert_not_called()

        self.assertEqual(res, "could not get Mozilla Firefox executable path")
        
    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "_get_executable_path", return_value="/path/to/firefox/executable")
    def test_create_window_failed_to_run(self, mock_get_executable_path, mock_sleep):
        self.firefox.dll.AU3_Run.return_value = 0
        res = self.firefox.create_window()
        
        mock_get_executable_path.assert_called_once()
        self.firefox.dll.AU3_Run.assert_called_once_with("/path/to/firefox/executable", "", 1)
        self.firefox.dll.AU3_WinExists.assert_not_called()
        self.firefox.dll.AU3_WinGetHandle.assert_not_called()
        self.firefox.dll.AU3_WinActivate.assert_not_called()
        self.firefox.dll.AU3_WinWaitActive.assert_not_called()
        self.firefox.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(res, "could not run Mozilla Firefox")
        
    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "_get_executable_path", return_value="/path/to/firefox/executable")
    def test_create_window_failed_window_didnt_exist(self, mock_get_executable_path, mock_sleep):
        self.firefox.dll.AU3_Run.return_value = 1
        self.firefox.dll.AU3_WinExists.return_value = 0

        res = self.firefox.create_window()
        
        mock_get_executable_path.assert_called_once()
        self.firefox.dll.AU3_Run.assert_called_once_with("/path/to/firefox/executable", "", 1)
        self.firefox.dll.AU3_WinExists.assert_called_once()
        self.firefox.dll.AU3_WinGetHandle.assert_not_called()
        self.firefox.dll.AU3_WinActivate.assert_not_called()
        self.firefox.dll.AU3_WinWaitActive.assert_not_called()
        self.firefox.dll.AU3_WinSetState.assert_not_called()

        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "Mozilla Firefox window didn't exist")
        
    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "_get_executable_path", return_value="/path/to/firefox/executable")
    def test_create_window_failed_to_get_handle(self, mock_get_executable_path, mock_sleep):
        self.firefox.dll.AU3_Run.return_value = 1
        self.firefox.dll.AU3_WinExists.return_value = 1
        self.firefox.dll.AU3_WinGetHandle.return_value = 0

        res = self.firefox.create_window()
        
        mock_get_executable_path.assert_called_once()
        self.firefox.dll.AU3_Run.assert_called_once_with("/path/to/firefox/executable", "", 1)
        self.firefox.dll.AU3_WinExists.assert_called_once()
        self.firefox.dll.AU3_WinGetHandle.assert_called_once()
        self.firefox.dll.AU3_WinActivate.assert_not_called()
        self.firefox.dll.AU3_WinWaitActive.assert_not_called()
        self.firefox.dll.AU3_WinSetState.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not get window handle")
        
    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "_get_executable_path", return_value="/path/to/firefox/executable")
    def test_create_window_failed_to_wait_window_activation(self, mock_get_executable_path, mock_sleep):
        self.firefox.dll.AU3_Run.return_value = 1
        self.firefox.dll.AU3_WinExists.return_value = 1
        self.firefox.dll.AU3_WinGetHandle.return_value = 1234567890
        self.firefox.dll.AU3_WinActivate.return_value = 1
        self.firefox.dll.AU3_WinWaitActive.return_value = 0

        res = self.firefox.create_window()
        
        mock_get_executable_path.assert_called_once()
        self.firefox.dll.AU3_Run.assert_called_once_with("/path/to/firefox/executable", "", 1)
        self.firefox.dll.AU3_WinExists.assert_called_once()
        self.firefox.dll.AU3_WinGetHandle.assert_called_once()
        self.firefox.dll.AU3_WinActivate.assert_called_once_with(self.firefox.window_info, "")
        self.firefox.dll.AU3_WinWaitActive.assert_called_once_with(self.firefox.window_info, "", 10)
        self.firefox.dll.AU3_WinSetState.assert_not_called()

        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not activate Mozilla Firefox")
        
    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "_get_executable_path", return_value="/path/to/firefox/executable")
    def test_create_window_failed_to_maximize_window(self, mock_get_executable_path, mock_sleep):
        self.firefox.dll.AU3_Run.return_value = 1
        self.firefox.dll.AU3_WinExists.return_value = 1
        self.firefox.dll.AU3_WinGetHandle.return_value = 1234567890
        self.firefox.dll.AU3_WinActivate.return_value = 1
        self.firefox.dll.AU3_WinWaitActive.return_value = 1
        self.firefox.dll.AU3_WinSetState.return_value = 0

        res = self.firefox.create_window()
        
        mock_get_executable_path.assert_called_once()
        self.firefox.dll.AU3_Run.assert_called_once_with("/path/to/firefox/executable", "", 1)
        self.firefox.dll.AU3_WinExists.assert_called_once()
        self.firefox.dll.AU3_WinGetHandle.assert_called_once()
        self.firefox.dll.AU3_WinActivate.assert_called_once_with(self.firefox.window_info, "")
        self.firefox.dll.AU3_WinWaitActive.assert_called_once_with(self.firefox.window_info, "", 10)
        self.firefox.dll.AU3_WinSetState.assert_called_once_with(self.firefox.window_info, "", 3)

        self.assertEqual(mock_sleep.call_count, 2)
        self.assertEqual(res, "could not maximize Mozilla Firefox window")
        
    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "check_existing_window", return_value=None)
    def test_create_tab_success(self, mock_check_existing_window, mock_sleep):
        self.firefox.dll.AU3_WinSetState.return_value = 1
        self.firefox.dll.AU3_Send.return_value = 1
        
        res = self.firefox.create_tab()
        
        mock_check_existing_window.assert_called_once()
        self.firefox.dll.AU3_WinSetState.assert_called_once_with(self.firefox.window_info, "", 3)
        self.firefox.dll.AU3_Send.assert_called_once_with("^t", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "check_existing_window", return_value=None)
    def test_create_tab_failed_to_maximize_window(self, mock_check_existing_window, mock_sleep):
        self.firefox.dll.AU3_WinSetState.return_value = 0
        
        res = self.firefox.create_tab()
        
        mock_check_existing_window.assert_called_once()
        self.firefox.dll.AU3_WinSetState.assert_called_once_with(self.firefox.window_info, "", 3)
        self.firefox.dll.AU3_Send.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not maximize Mozilla Firefox window")

    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "check_existing_window", return_value=None)
    def test_create_tab_failed_to_send_keys(self, mock_check_existing_window, mock_sleep):
        self.firefox.dll.AU3_WinSetState.return_value = 1
        self.firefox.dll.AU3_Send.return_value = 0
        
        res = self.firefox.create_tab()
        
        mock_check_existing_window.assert_called_once()
        self.firefox.dll.AU3_WinSetState.assert_called_once_with(self.firefox.window_info, "", 3)
        self.firefox.dll.AU3_Send.assert_called_once_with("^t", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not send keys to create new tab")
        
    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "check_existing_window", return_value=None)
    def test_browse_success(self, mock_check_existing_window, mock_sleep):
        url = "www.python.org"
        self.firefox.dll.AU3_WinActive.return_value = 1
        self.firefox.dll.AU3_Send.return_value = 1
        
        res = self.firefox.browse(url=url)
        
        mock_check_existing_window.assert_called_once()
        self.firefox.dll.AU3_WinActive.assert_has_calls([call(self.firefox.window_info, "")] * len(url))
        self.firefox.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in url],
            call("{ENTER}", 0)
        ], 0)
        
        self.assertEqual(self.firefox.dll.AU3_WinActive.call_count, len(url))
        self.assertEqual(self.firefox.dll.AU3_Send.call_count, len(url) + 1)
        self.assertEqual(mock_sleep.call_count, len(url) + 2)
        self.assertIsNone(res)
    
    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "check_existing_window", return_value=None)
    def test_browse_failed_window_inactive(self, mock_check_existing_window, mock_sleep):
        url = "www.python.org"
        self.firefox.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        self.firefox.dll.AU3_Send.return_value = 1
        
        res = self.firefox.browse(url=url)
        
        mock_check_existing_window.assert_called_once()
        self.firefox.dll.AU3_WinActive.assert_has_calls([call(self.firefox.window_info, "")] * 6)
        self.firefox.dll.AU3_Send.assert_has_calls([*[call(char, 1) for char in "www.p"]])
        
        self.assertEqual(self.firefox.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.firefox.dll.AU3_Send.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "Mozilla Firefox window is inactive")
        
    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "check_existing_window", return_value=None)
    def test_browse_failed_to_send_letter(self, mock_check_existing_window, mock_sleep):
        url = "www.python.org"
        self.firefox.dll.AU3_WinActive.return_value = 1
        def send_side_effect(char, _):
            return False if char == "y" else True
        self.firefox.dll.AU3_Send.side_effect = send_side_effect
        
        res = self.firefox.browse(url=url)
        
        mock_check_existing_window.assert_called_once()
        self.firefox.dll.AU3_WinActive.assert_has_calls([call(self.firefox.window_info, "")] * 6)
        self.firefox.dll.AU3_Send.assert_has_calls([*[call(char, 1) for char in "www.py"]])
        
        self.assertEqual(self.firefox.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.firefox.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "could not send y to Mozilla Firefox")
        
    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "check_existing_window", return_value=None)
    def test_browse_failed_to_send_enter(self, mock_check_existing_window, mock_sleep):
        url = "www.python.org"
        self.firefox.dll.AU3_WinActive.return_value = 1
        self.firefox.dll.AU3_Send.side_effect = [1] * len(url) + [0]
        
        res = self.firefox.browse(url=url)
        
        mock_check_existing_window.assert_called_once()
        self.firefox.dll.AU3_WinActive.assert_has_calls([call(self.firefox.window_info, "")] * len(url))
        self.firefox.dll.AU3_Send.assert_has_calls([
            *[call(char, 1) for char in url],
            call("{ENTER}", 0)
        ])
        
        self.assertEqual(self.firefox.dll.AU3_WinActive.call_count, len(url))
        self.assertEqual(self.firefox.dll.AU3_Send.call_count, len(url) + 1)
        self.assertEqual(mock_sleep.call_count, len(url) + 2)
        self.assertEqual(res, "could not send Enter key")
        
    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "check_existing_window", return_value=None)
    def test_scroll_success(self, mock_check_existing_window, mock_sleep):
        direction = "down"
        clicks = 30
        scroll_delay = 0.07
        self.firefox.dll.AU3_WinActive.return_value = 1
        self.firefox.dll.AU3_MouseWheel.return_value = 1
        
        res = self.firefox.scroll(direction=direction, clicks=clicks, scroll_delay=scroll_delay)
        
        mock_check_existing_window.assert_called_once()
        self.firefox.dll.AU3_WinActive.assert_has_calls([call(self.firefox.window_info, "")] * clicks)
        self.firefox.dll.AU3_MouseWheel.assert_has_calls([call(direction, 1)] * clicks)
        
        self.assertEqual(self.firefox.dll.AU3_WinActive.call_count, clicks)
        self.assertEqual(self.firefox.dll.AU3_MouseWheel.call_count, clicks)
        self.assertEqual(mock_sleep.call_count, clicks + 1)
        self.assertIsNone(res)
        
    def test_scroll_failed_invalid_scroll_direction(self):
        direction = "abc"
        clicks = 30
        scroll_delay = 0.07
        
        res = self.firefox.scroll(direction, clicks, scroll_delay)
        
        self.assertEqual(res, "invalid scroll direction")
        
    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "check_existing_window", return_value=None)
    def test_scroll_failed_window_inactive(self, mock_check_existing_window, mock_sleep):
        direction = "down"
        clicks = 30
        scroll_delay = 0.07
        self.firefox.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        self.firefox.dll.AU3_MouseWheel.return_value = 1
        
        res = self.firefox.scroll(direction=direction, clicks=clicks, scroll_delay=scroll_delay)
        
        mock_check_existing_window.assert_called_once()
        self.firefox.dll.AU3_WinActive.assert_has_calls([call(self.firefox.window_info, "")] * 6)
        self.firefox.dll.AU3_MouseWheel.assert_has_calls([call(direction, 1)] * 5)
        
        self.assertEqual(self.firefox.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.firefox.dll.AU3_MouseWheel.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "Mozilla Firefox window is inactive")

    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "check_existing_window", return_value=None)
    def test_scroll_failed_to_move_mouse_wheel(self, mock_check_existing_window, mock_sleep):
        direction = "down"
        clicks = 30
        scroll_delay = 0.07
        self.firefox.dll.AU3_WinActive.return_value = 1
        self.firefox.dll.AU3_MouseWheel.side_effect = [1] * 5 + [0]
        
        res = self.firefox.scroll(direction=direction, clicks=clicks, scroll_delay=scroll_delay)
        
        mock_check_existing_window.assert_called_once()
        self.firefox.dll.AU3_WinActive.assert_has_calls([call(self.firefox.window_info, "")] * 6)
        self.firefox.dll.AU3_MouseWheel.assert_has_calls([call(direction, 1)] * 6)
        
        self.assertEqual(self.firefox.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.firefox.dll.AU3_MouseWheel.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 7)
        self.assertEqual(res, "could not scroll mouse wheel")
        
    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "check_existing_window", return_value=None)
    def test_zoom_in_success(self, mock_check_existing_window, mock_sleep):
        count = 10
        self.firefox.dll.AU3_WinActive.return_value = 1
        self.firefox.dll.AU3_Send.return_value = 1
        
        res = self.firefox.zoom_in(count=count)
        
        mock_check_existing_window.assert_called_once()
        self.firefox.dll.AU3_WinActive.assert_has_calls([call(self.firefox.window_info, "")] * count)
        self.firefox.dll.AU3_Send.assert_has_calls([call("^{+}", 0)] * count)
        
        self.assertEqual(self.firefox.dll.AU3_WinActive.call_count, count)
        self.assertEqual(self.firefox.dll.AU3_Send.call_count, count)
        self.assertEqual(mock_sleep.call_count, count + 1)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "check_existing_window", return_value=None)
    def test_zoom_in_failed_window_inactive(self, mock_check_existing_window, mock_sleep):
        count = 10
        self.firefox.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        self.firefox.dll.AU3_Send.return_value = 1
        
        res = self.firefox.zoom_in(count=count)
        
        mock_check_existing_window.assert_called_once()
        self.firefox.dll.AU3_WinActive.assert_has_calls([call(self.firefox.window_info, "")] * 6)
        self.firefox.dll.AU3_Send.assert_has_calls([call("^{+}", 0)] * 5)
        
        self.assertEqual(self.firefox.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.firefox.dll.AU3_Send.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "Mozilla Firefox window is inactive")

    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "check_existing_window", return_value=None)
    def test_zoom_in_failed_to_send_keys(self, mock_check_existing_window, mock_sleep):
        count = 10
        self.firefox.dll.AU3_WinActive.return_value = 1
        self.firefox.dll.AU3_Send.side_effect = [1] * 5 + [0]
        
        res = self.firefox.zoom_in(count=count)
        
        mock_check_existing_window.assert_called_once()
        self.firefox.dll.AU3_WinActive.assert_has_calls([call(self.firefox.window_info, "")] * 6)
        self.firefox.dll.AU3_Send.assert_has_calls([call("^{+}", 0)] * 6)
        
        self.assertEqual(self.firefox.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.firefox.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "could not send keys to zoom in")
        
    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "check_existing_window", return_value=None)
    def test_zoom_out_success(self, mock_check_existing_window, mock_sleep):
        count = 10
        self.firefox.dll.AU3_WinActive.return_value = 1
        self.firefox.dll.AU3_Send.return_value = 1
        
        res = self.firefox.zoom_out(count=count)
        
        mock_check_existing_window.assert_called_once()
        self.firefox.dll.AU3_WinActive.assert_has_calls([call(self.firefox.window_info, "")] * count)
        self.firefox.dll.AU3_Send.assert_has_calls([call("^{-}", 0)] * count)
        
        self.assertEqual(self.firefox.dll.AU3_WinActive.call_count, count)
        self.assertEqual(self.firefox.dll.AU3_Send.call_count, count)
        self.assertEqual(mock_sleep.call_count, count + 1)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "check_existing_window", return_value=None)
    def test_zoom_out_failed_window_inactive(self, mock_check_existing_window, mock_sleep):
        count = 10
        self.firefox.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        self.firefox.dll.AU3_Send.return_value = 1
        
        res = self.firefox.zoom_out(count=count)
        
        mock_check_existing_window.assert_called_once()
        self.firefox.dll.AU3_WinActive.assert_has_calls([call(self.firefox.window_info, "")] * 6)
        self.firefox.dll.AU3_Send.assert_has_calls([call("^{-}", 0)] * 5)
        
        self.assertEqual(self.firefox.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.firefox.dll.AU3_Send.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "Mozilla Firefox window is inactive")

    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "check_existing_window", return_value=None)
    def test_zoom_out_failed_to_send_keys(self, mock_check_existing_window, mock_sleep):
        count = 10
        self.firefox.dll.AU3_WinActive.return_value = 1
        self.firefox.dll.AU3_Send.side_effect = [1] * 5 + [0]
        
        res = self.firefox.zoom_out(count=count)
        
        mock_check_existing_window.assert_called_once()
        self.firefox.dll.AU3_WinActive.assert_has_calls([call(self.firefox.window_info, "")] * 6)
        self.firefox.dll.AU3_Send.assert_has_calls([call("^{-}", 0)] * 6)
        
        self.assertEqual(self.firefox.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.firefox.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "could not send keys to zoom out")
        
    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "check_existing_window", return_value=None)
    def test_previous_page_success(self, mock_check_existing_window, mock_sleep):
        count = 10
        self.firefox.dll.AU3_WinActive.return_value = 1
        self.firefox.dll.AU3_Send.return_value = 1
        
        res = self.firefox.previous_page(count=count)
        
        mock_check_existing_window.assert_called_once()
        self.firefox.dll.AU3_WinActive.assert_has_calls([call(self.firefox.window_info, "")] * count)
        self.firefox.dll.AU3_Send.assert_has_calls([call("!{LEFT}", 0)] * count)
        
        self.assertEqual(self.firefox.dll.AU3_WinActive.call_count, count)
        self.assertEqual(self.firefox.dll.AU3_Send.call_count, count)
        self.assertEqual(mock_sleep.call_count, count + 1)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "check_existing_window", return_value=None)
    def test_previous_page_failed_window_inactive(self, mock_check_existing_window, mock_sleep):
        count = 10
        self.firefox.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        self.firefox.dll.AU3_Send.return_value = 1
        
        res = self.firefox.previous_page(count=count)
        
        mock_check_existing_window.assert_called_once()
        self.firefox.dll.AU3_WinActive.assert_has_calls([call(self.firefox.window_info, "")] * 6)
        self.firefox.dll.AU3_Send.assert_has_calls([call("!{LEFT}", 0)] * 5)
        
        self.assertEqual(self.firefox.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.firefox.dll.AU3_Send.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "Mozilla Firefox window is inactive")

    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "check_existing_window", return_value=None)
    def test_previous_page_failed_to_send_keys(self, mock_check_existing_window, mock_sleep):
        count = 10
        self.firefox.dll.AU3_WinActive.return_value = 1
        self.firefox.dll.AU3_Send.side_effect = [1] * 5 + [0]
        
        res = self.firefox.previous_page(count=count)
        
        mock_check_existing_window.assert_called_once()
        self.firefox.dll.AU3_WinActive.assert_has_calls([call(self.firefox.window_info, "")] * 6)
        self.firefox.dll.AU3_Send.assert_has_calls([call("!{LEFT}", 0)] * 6)
        
        self.assertEqual(self.firefox.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.firefox.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "could not send keys to go to previous page")
   
    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "check_existing_window", return_value=None)
    def test_next_page_success(self, mock_check_existing_window, mock_sleep):
        count = 10
        self.firefox.dll.AU3_WinActive.return_value = 1
        self.firefox.dll.AU3_Send.return_value = 1
        
        res = self.firefox.next_page(count=count)
        
        mock_check_existing_window.assert_called_once()
        self.firefox.dll.AU3_WinActive.assert_has_calls([call(self.firefox.window_info, "")] * count)
        self.firefox.dll.AU3_Send.assert_has_calls([call("!{RIGHT}", 0)] * count)
        
        self.assertEqual(self.firefox.dll.AU3_WinActive.call_count, count)
        self.assertEqual(self.firefox.dll.AU3_Send.call_count, count)
        self.assertEqual(mock_sleep.call_count, count + 1)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "check_existing_window", return_value=None)
    def test_next_page_failed_window_inactive(self, mock_check_existing_window, mock_sleep):
        count = 10
        self.firefox.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        self.firefox.dll.AU3_Send.return_value = 1
        
        res = self.firefox.next_page(count=count)
        
        mock_check_existing_window.assert_called_once()
        self.firefox.dll.AU3_WinActive.assert_has_calls([call(self.firefox.window_info, "")] * 6)
        self.firefox.dll.AU3_Send.assert_has_calls([call("!{RIGHT}", 0)] * 5)
        
        self.assertEqual(self.firefox.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.firefox.dll.AU3_Send.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "Mozilla Firefox window is inactive")

    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "check_existing_window", return_value=None)
    def test_next_page_failed_to_send_keys(self, mock_check_existing_window, mock_sleep):
        count = 10
        self.firefox.dll.AU3_WinActive.return_value = 1
        self.firefox.dll.AU3_Send.side_effect = [1] * 5 + [0]
        
        res = self.firefox.next_page(count=count)
        
        mock_check_existing_window.assert_called_once()
        self.firefox.dll.AU3_WinActive.assert_has_calls([call(self.firefox.window_info, "")] * 6)
        self.firefox.dll.AU3_Send.assert_has_calls([call("!{RIGHT}", 0)] * 6)
        
        self.assertEqual(self.firefox.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.firefox.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "could not send keys to go to next page")
    
    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "check_existing_window", return_value=None)
    def test_toggle_fullscreen_success(self, mock_check_existing_window, mock_sleep):
        count = 10
        self.firefox.dll.AU3_WinActive.return_value = 1
        self.firefox.dll.AU3_Send.return_value = 1
        
        res = self.firefox.toggle_fullscreen(count=count)
        
        mock_check_existing_window.assert_called_once()
        self.firefox.dll.AU3_WinActive.assert_has_calls([call(self.firefox.window_info, "")] * count)
        self.firefox.dll.AU3_Send.assert_has_calls([call("{F11}", 0)] * count)
        
        self.assertEqual(self.firefox.dll.AU3_WinActive.call_count, count)
        self.assertEqual(self.firefox.dll.AU3_Send.call_count, count)
        self.assertEqual(mock_sleep.call_count, count + 1)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "check_existing_window", return_value=None)
    def test_toggle_fullscreen_failed_window_inactive(self, mock_check_existing_window, mock_sleep):
        count = 10
        self.firefox.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        self.firefox.dll.AU3_Send.return_value = 1
        
        res = self.firefox.toggle_fullscreen(count=count)
        
        mock_check_existing_window.assert_called_once()
        self.firefox.dll.AU3_WinActive.assert_has_calls([call(self.firefox.window_info, "")] * 6)
        self.firefox.dll.AU3_Send.assert_has_calls([call("{F11}", 0)] * 5)
        
        self.assertEqual(self.firefox.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.firefox.dll.AU3_Send.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "Mozilla Firefox window is inactive")

    @patch("time.sleep", return_value=None)
    @patch.object(MozillaFirefox, "check_existing_window", return_value=None)
    def test_toggle_fullscreen_failed_to_send_keys(self, mock_check_existing_window, mock_sleep):
        count = 10
        self.firefox.dll.AU3_WinActive.return_value = 1
        self.firefox.dll.AU3_Send.side_effect = [1] * 5 + [0]
        
        res = self.firefox.toggle_fullscreen(count=count)
        
        mock_check_existing_window.assert_called_once()
        self.firefox.dll.AU3_WinActive.assert_has_calls([call(self.firefox.window_info, "")] * 6)
        self.firefox.dll.AU3_Send.assert_has_calls([call("{F11}", 0)] * 6)
        
        self.assertEqual(self.firefox.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.firefox.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "could not send key to toggle fullscreen")