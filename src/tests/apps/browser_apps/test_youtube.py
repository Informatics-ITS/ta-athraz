import unittest
from unittest.mock import patch, call, MagicMock
from activities.apps.browser_apps.youtube import YouTube

class TestYouTube(unittest.TestCase):
    def setUp(self):
        self.youtube = YouTube(browser=MagicMock())
        self.youtube.dll = MagicMock()
    
    def test_open_success_with_create_window(self):
        self.youtube.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.youtube.browser.create_window.return_value = None
        self.youtube.browser.browse.return_value = None

        res = self.youtube.open()

        self.youtube.browser.check_existing_window.assert_called_once()
        self.youtube.browser.create_window.assert_called_once()
        self.youtube.browser.create_tab.assert_not_called()
        self.youtube.browser.browse.assert_called_once_with("https://youtube.com")
        self.assertIsNone(res)
        
    def test_open_success_with_create_tab(self):
        self.youtube.browser.check_existing_window.return_value = None
        self.youtube.browser.create_tab.return_value = None
        self.youtube.browser.browse.return_value = None

        res = self.youtube.open()

        self.youtube.browser.check_existing_window.assert_called_once()
        self.youtube.browser.create_window.assert_not_called()
        self.youtube.browser.create_tab.assert_called_once()
        self.youtube.browser.browse.assert_called_once_with("https://youtube.com")
        self.assertIsNone(res)
    
    @patch("time.sleep", return_value=None)
    def test_search_success(self, mock_sleep):
        text = "python tutorial"
        self.youtube.browser.check_existing_window.return_value = None
        self.youtube.dll.AU3_Send.return_value = 1
        self.youtube.dll.AU3_WinActive.return_value = 1
        
        res = self.youtube.search(text=text)
        self.youtube.browser.check_existing_window.assert_called_once()
        self.youtube.dll.AU3_Send.assert_has_calls([
            call("/", 0),
            call("^a{BACKSPACE}", 0),
            *[call(char, 1) for char in text],
            call("{ENTER}", 0)
        ])
        self.youtube.dll.AU3_WinActive.assert_called_with(self.youtube.browser.window_info, "")
        
        self.assertEqual(self.youtube.dll.AU3_WinActive.call_count, len(text))
        self.assertEqual(self.youtube.dll.AU3_Send.call_count, len(text) + 3)
        self.assertEqual(mock_sleep.call_count, len(text) + 4)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    def test_search_failed_to_send_search_box_keys(self, mock_sleep):
        text = "python tutorial"
        self.youtube.browser.check_existing_window.return_value = None
        self.youtube.dll.AU3_Send.return_value = 0
        
        res = self.youtube.search(text=text)
        self.youtube.browser.check_existing_window.assert_called_once()
        self.youtube.dll.AU3_Send.assert_called_once_with("/", 0)
        self.youtube.dll.AU3_WinActive.assert_not_called()
        
        self.assertEqual(self.youtube.dll.AU3_Send.call_count, 1)
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not send keys to go to search box")
        
    @patch("time.sleep", return_value=None)
    def test_search_failed_to_clear_search_box(self, mock_sleep):
        text = "python tutorial"
        self.youtube.browser.check_existing_window.return_value = None
        self.youtube.dll.AU3_Send.side_effect = [1, 0]
        
        res = self.youtube.search(text=text)
        self.youtube.browser.check_existing_window.assert_called_once()
        self.youtube.dll.AU3_Send.assert_has_calls([
            call("/", 0),
            call("^a{BACKSPACE}", 0)
        ])
        self.youtube.dll.AU3_WinActive.assert_not_called()
        
        self.assertEqual(self.youtube.dll.AU3_Send.call_count, 2)
        self.assertEqual(mock_sleep.call_count, 2)
        self.assertEqual(res, "could not send keys to clear search box")
        
    @patch("time.sleep", return_value=None)
    def test_search_failed_window_inactive(self, mock_sleep):
        text = "python tutorial"
        self.youtube.browser.check_existing_window.return_value = None
        self.youtube.dll.AU3_Send.return_value = 1
        self.youtube.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        
        res = self.youtube.search(text=text)
        self.youtube.browser.check_existing_window.assert_called_once()
        self.youtube.dll.AU3_Send.assert_has_calls([
            call("/", 0),
            call("^a{BACKSPACE}", 0),
            *[call(char, 1) for char in "pytho"]
        ])
        self.youtube.dll.AU3_WinActive.assert_called_with(self.youtube.browser.window_info, "")
        
        self.assertEqual(self.youtube.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.youtube.dll.AU3_Send.call_count, 7)
        self.assertEqual(mock_sleep.call_count, 8)
        self.assertEqual(res, "browser window is inactive while filling YouTube search box")
        
    @patch("time.sleep", return_value=None)
    def test_search_failed_to_send_letter(self, mock_sleep):
        text = "python tutorial"
        self.youtube.browser.check_existing_window.return_value = None
        self.youtube.dll.AU3_Send.side_effect = [1] * 5 + [0]
        self.youtube.dll.AU3_WinActive.return_value = 1
        
        res = self.youtube.search(text=text)
        self.youtube.browser.check_existing_window.assert_called_once()
        self.youtube.dll.AU3_Send.assert_has_calls([
            call("/", 0),
            call("^a{BACKSPACE}", 0),
            *[call(char, 1) for char in "pyth"]
        ])
        self.youtube.dll.AU3_WinActive.assert_called_with(self.youtube.browser.window_info, "")
        
        self.assertEqual(self.youtube.dll.AU3_WinActive.call_count, 4)
        self.assertEqual(self.youtube.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "could not send h to YouTube search box")
    
    @patch("time.sleep", return_value=None)
    def test_search_failed_to_send_enter(self, mock_sleep):
        text = "python tutorial"
        self.youtube.browser.check_existing_window.return_value = None
        self.youtube.dll.AU3_Send.side_effect = [1] * (len(text) + 2) + [0]
        self.youtube.dll.AU3_WinActive.return_value = 1
        
        res = self.youtube.search(text=text)
        self.youtube.browser.check_existing_window.assert_called_once()
        self.youtube.dll.AU3_Send.assert_has_calls([
            call("/", 0),
            call("^a{BACKSPACE}", 0),
            *[call(char, 1) for char in text],
            call("{ENTER}", 0)
        ])
        self.youtube.dll.AU3_WinActive.assert_called_with(self.youtube.browser.window_info, "")
        
        self.assertEqual(self.youtube.dll.AU3_WinActive.call_count, len(text))
        self.assertEqual(self.youtube.dll.AU3_Send.call_count, len(text) + 3)
        self.assertEqual(mock_sleep.call_count, len(text) + 4)
        self.assertEqual(res, "could not send enter key to search")
    
    @patch("time.sleep", return_value=None)
    def test_toggle_pause_success(self, mock_sleep):
        self.youtube.browser.check_existing_window.return_value = None
        self.youtube.dll.AU3_Send.return_value = 1
        
        res = self.youtube.toggle_pause()
        self.youtube.browser.check_existing_window.assert_called_once()
        self.youtube.dll.AU3_Send.assert_called_once_with("{SPACE}", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    def test_toggle_pause_failed_to_send_key(self, mock_sleep):
        self.youtube.browser.check_existing_window.return_value = None
        self.youtube.dll.AU3_Send.return_value = 0
        
        res = self.youtube.toggle_pause()
        self.youtube.browser.check_existing_window.assert_called_once()
        self.youtube.dll.AU3_Send.assert_called_once_with("{SPACE}", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not send keys to toggle pause")
        
    @patch("time.sleep", return_value=None)
    def test_toggle_mute_success(self, mock_sleep):
        self.youtube.browser.check_existing_window.return_value = None
        self.youtube.dll.AU3_Send.return_value = 1
        
        res = self.youtube.toggle_mute()
        self.youtube.browser.check_existing_window.assert_called_once()
        self.youtube.dll.AU3_Send.assert_called_once_with("m", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    def test_toggle_mute_failed_to_send_key(self, mock_sleep):
        self.youtube.browser.check_existing_window.return_value = None
        self.youtube.dll.AU3_Send.return_value = 0
        
        res = self.youtube.toggle_mute()
        self.youtube.browser.check_existing_window.assert_called_once()
        self.youtube.dll.AU3_Send.assert_called_once_with("m", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not send keys to toggle mute")
        
    @patch("time.sleep", return_value=None)
    def test_toggle_subtitle_success(self, mock_sleep):
        self.youtube.browser.check_existing_window.return_value = None
        self.youtube.dll.AU3_Send.return_value = 1
        
        res = self.youtube.toggle_subtitle()
        self.youtube.browser.check_existing_window.assert_called_once()
        self.youtube.dll.AU3_Send.assert_called_once_with("c", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    def test_toggle_subtitle_failed_to_send_key(self, mock_sleep):
        self.youtube.browser.check_existing_window.return_value = None
        self.youtube.dll.AU3_Send.return_value = 0
        
        res = self.youtube.toggle_subtitle()
        self.youtube.browser.check_existing_window.assert_called_once()
        self.youtube.dll.AU3_Send.assert_called_once_with("c", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not send keys to toggle subtitle")
        
    @patch("time.sleep", return_value=None)
    def test_toggle_fullscreen_success(self, mock_sleep):
        self.youtube.browser.check_existing_window.return_value = None
        self.youtube.dll.AU3_Send.return_value = 1
        
        res = self.youtube.toggle_fullscreen()
        self.youtube.browser.check_existing_window.assert_called_once()
        self.youtube.dll.AU3_Send.assert_called_once_with("f", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    def test_toggle_fullscreen_failed_to_send_key(self, mock_sleep):
        self.youtube.browser.check_existing_window.return_value = None
        self.youtube.dll.AU3_Send.return_value = 0
        
        res = self.youtube.toggle_fullscreen()
        self.youtube.browser.check_existing_window.assert_called_once()
        self.youtube.dll.AU3_Send.assert_called_once_with("f", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not send keys to toggle fullscreen")
        
    @patch("time.sleep", return_value=None)
    def test_toggle_cinema_mode_success(self, mock_sleep):
        self.youtube.browser.check_existing_window.return_value = None
        self.youtube.dll.AU3_Send.return_value = 1
        
        res = self.youtube.toggle_cinema_mode()
        self.youtube.browser.check_existing_window.assert_called_once()
        self.youtube.dll.AU3_Send.assert_called_once_with("t", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    def test_toggle_cinema_mode_failed_to_send_key(self, mock_sleep):
        self.youtube.browser.check_existing_window.return_value = None
        self.youtube.dll.AU3_Send.return_value = 0
        
        res = self.youtube.toggle_cinema_mode()
        self.youtube.browser.check_existing_window.assert_called_once()
        self.youtube.dll.AU3_Send.assert_called_once_with("t", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not send keys to toggle cinema mode")
        
    @patch("time.sleep", return_value=None)
    def test_next_video_success(self, mock_sleep):
        count = 10
        self.youtube.browser.check_existing_window.return_value = None
        self.youtube.dll.AU3_WinActive.return_value = 1
        self.youtube.dll.AU3_Send.return_value = 1
        
        res = self.youtube.next_video(count=count)
        self.youtube.browser.check_existing_window.assert_called_once()
        self.youtube.dll.AU3_WinActive.assert_has_calls([call(self.youtube.browser.window_info, "")] * count)
        self.youtube.dll.AU3_Send.assert_has_calls([call("+n", 0)] * count)
        
        self.assertEqual(self.youtube.dll.AU3_WinActive.call_count, count)
        self.assertEqual(self.youtube.dll.AU3_Send.call_count, count)
        self.assertEqual(mock_sleep.call_count, count + 1)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    def test_next_video_failed_window_inactive(self, mock_sleep):
        count = 10
        self.youtube.browser.check_existing_window.return_value = None
        self.youtube.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        self.youtube.dll.AU3_Send.return_value = 1
        
        res = self.youtube.next_video(count=count)
        self.youtube.browser.check_existing_window.assert_called_once()
        self.youtube.dll.AU3_WinActive.assert_has_calls([call(self.youtube.browser.window_info, "")] * 6)
        self.youtube.dll.AU3_Send.assert_has_calls([call("+n", 0)] * 5)
        
        self.assertEqual(self.youtube.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.youtube.dll.AU3_Send.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "browser window is inactive")
        
    @patch("time.sleep", return_value=None)
    def test_next_video_failed_to_send_keys(self, mock_sleep):
        count = 10
        self.youtube.browser.check_existing_window.return_value = None
        self.youtube.dll.AU3_WinActive.return_value = 1
        self.youtube.dll.AU3_Send.side_effect = [1] * 5 + [0]
        
        res = self.youtube.next_video(count=count)
        self.youtube.browser.check_existing_window.assert_called_once()
        self.youtube.dll.AU3_WinActive.assert_has_calls([call(self.youtube.browser.window_info, "")] * 6)
        self.youtube.dll.AU3_Send.assert_has_calls([call("+n", 0)] * 6)
        
        self.assertEqual(self.youtube.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.youtube.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "could not send keys to move to next video")
        
    @patch("time.sleep", return_value=None)
    def test_previous_video_success(self, mock_sleep):
        count = 10
        self.youtube.browser.check_existing_window.return_value = None
        self.youtube.dll.AU3_WinActive.return_value = 1
        self.youtube.dll.AU3_Send.return_value = 1
        
        res = self.youtube.previous_video(count=count)
        self.youtube.browser.check_existing_window.assert_called_once()
        self.youtube.dll.AU3_WinActive.assert_has_calls([call(self.youtube.browser.window_info, "")] * count)
        self.youtube.dll.AU3_Send.assert_has_calls([call("+p", 0)] * count)
        
        self.assertEqual(self.youtube.dll.AU3_WinActive.call_count, count)
        self.assertEqual(self.youtube.dll.AU3_Send.call_count, count)
        self.assertEqual(mock_sleep.call_count, count + 1)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    def test_previous_video_failed_window_inactive(self, mock_sleep):
        count = 10
        self.youtube.browser.check_existing_window.return_value = None
        self.youtube.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        self.youtube.dll.AU3_Send.return_value = 1
        
        res = self.youtube.previous_video(count=count)
        self.youtube.browser.check_existing_window.assert_called_once()
        self.youtube.dll.AU3_WinActive.assert_has_calls([call(self.youtube.browser.window_info, "")] * 6)
        self.youtube.dll.AU3_Send.assert_has_calls([call("+p", 0)] * 5)
        
        self.assertEqual(self.youtube.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.youtube.dll.AU3_Send.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "browser window is inactive")
        
    @patch("time.sleep", return_value=None)
    def test_previous_video_failed_to_send_keys(self, mock_sleep):
        count = 10
        self.youtube.browser.check_existing_window.return_value = None
        self.youtube.dll.AU3_WinActive.return_value = 1
        self.youtube.dll.AU3_Send.side_effect = [1] * 5 + [0]
        
        res = self.youtube.previous_video(count=count)
        self.youtube.browser.check_existing_window.assert_called_once()
        self.youtube.dll.AU3_WinActive.assert_has_calls([call(self.youtube.browser.window_info, "")] * 6)
        self.youtube.dll.AU3_Send.assert_has_calls([call("+p", 0)] * 6)
        
        self.assertEqual(self.youtube.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.youtube.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "could not send keys to move to previous video")