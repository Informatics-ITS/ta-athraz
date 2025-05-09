import unittest
from unittest.mock import patch, call, MagicMock
from activities.apps.browser_apps.gmail import Gmail

class TestGmail(unittest.TestCase):
    def setUp(self):
        self.gmail = Gmail(browser=MagicMock())
        self.gmail.dll = MagicMock()
        
    def test_open_success_with_create_window(self):
        self.gmail.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gmail.browser.create_window.return_value = None
        self.gmail.browser.browse.return_value = None

        res = self.gmail.open()

        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.browser.create_window.assert_called_once()
        self.gmail.browser.create_tab.assert_not_called()
        self.gmail.browser.browse.assert_called_once_with("https://mail.google.com")
        self.assertIsNone(res)
        
    def test_open_success_with_create_tab(self):
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.browser.create_tab.return_value = None
        self.gmail.browser.browse.return_value = None

        res = self.gmail.open()

        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.browser.create_window.assert_not_called()
        self.gmail.browser.create_tab.assert_called_once()
        self.gmail.browser.browse.assert_called_once_with("https://mail.google.com")
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    def test_open_email_success(self, mock_sleep):
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.return_value = 1
        
        res = self.gmail.open_email()
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_called_once_with("{ENTER}", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    def test_open_email_failed_to_send_key(self, mock_sleep):
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.return_value = 0
        
        res = self.gmail.open_email()
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_called_once_with("{ENTER}", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not send enter key to open email")
        
    @patch("time.sleep", return_value=None)
    def test_next_email_success(self, mock_sleep):
        count = 10
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_WinActive.return_value = 1
        self.gmail.dll.AU3_Send.return_value = 1
        
        res = self.gmail.next_email(count=count)
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_WinActive.assert_has_calls([call(self.gmail.browser.window_info, "")] * count)
        self.gmail.dll.AU3_Send.assert_has_calls([call("j", 0)] * count)
        
        self.assertEqual(self.gmail.dll.AU3_WinActive.call_count, count)
        self.assertEqual(self.gmail.dll.AU3_Send.call_count, count)
        self.assertEqual(mock_sleep.call_count, count + 1)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    def test_next_email_failed_window_inactive(self, mock_sleep):
        count = 10
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        self.gmail.dll.AU3_Send.return_value = 1
        
        res = self.gmail.next_email(count=count)
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_WinActive.assert_has_calls([call(self.gmail.browser.window_info, "")] * 6)
        self.gmail.dll.AU3_Send.assert_has_calls([call("j", 0)] * 5)
        
        self.assertEqual(self.gmail.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.gmail.dll.AU3_Send.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "browser window is inactive")
        
    @patch("time.sleep", return_value=None)
    def test_next_email_failed_to_send_keys(self, mock_sleep):
        count = 10
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_WinActive.return_value = 1
        self.gmail.dll.AU3_Send.side_effect = [1] * 5 + [0]
        
        res = self.gmail.next_email(count=count)
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_WinActive.assert_has_calls([call(self.gmail.browser.window_info, "")] * 6)
        self.gmail.dll.AU3_Send.assert_has_calls([call("j", 0)] * 6)
        
        self.assertEqual(self.gmail.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.gmail.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "could not send key to go to next email")

    @patch("time.sleep", return_value=None)
    def test_previous_email_success(self, mock_sleep):
        count = 10
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_WinActive.return_value = 1
        self.gmail.dll.AU3_Send.return_value = 1
        
        res = self.gmail.previous_email(count=count)
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_WinActive.assert_has_calls([call(self.gmail.browser.window_info, "")] * count)
        self.gmail.dll.AU3_Send.assert_has_calls([call("k", 0)] * count)
        
        self.assertEqual(self.gmail.dll.AU3_WinActive.call_count, count)
        self.assertEqual(self.gmail.dll.AU3_Send.call_count, count)
        self.assertEqual(mock_sleep.call_count, count + 1)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    def test_previous_email_failed_window_inactive(self, mock_sleep):
        count = 10
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        self.gmail.dll.AU3_Send.return_value = 1
        
        res = self.gmail.previous_email(count=count)
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_WinActive.assert_has_calls([call(self.gmail.browser.window_info, "")] * 6)
        self.gmail.dll.AU3_Send.assert_has_calls([call("k", 0)] * 5)
        
        self.assertEqual(self.gmail.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.gmail.dll.AU3_Send.call_count, 5)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "browser window is inactive")
        
    @patch("time.sleep", return_value=None)
    def test_previous_email_failed_to_send_keys(self, mock_sleep):
        count = 10
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_WinActive.return_value = 1
        self.gmail.dll.AU3_Send.side_effect = [1] * 5 + [0]
        
        res = self.gmail.previous_email(count=count)
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_WinActive.assert_has_calls([call(self.gmail.browser.window_info, "")] * 6)
        self.gmail.dll.AU3_Send.assert_has_calls([call("k", 0)] * 6)
        
        self.assertEqual(self.gmail.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.gmail.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 6)
        self.assertEqual(res, "could not send key to go to previous email")
        
    @patch("time.sleep", return_value=None)
    def test_compose_email_success(self, mock_sleep):
        recipients = ["example1@mail.com", "example2@mail.com", "example3@mail.com"]
        subject = "Example Subject Gmail"
        body = "example body gmail"
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.return_value = 1
        self.gmail.dll.AU3_WinActive.return_value = 1
        
        res = self.gmail.compose_email(recipients=recipients, subject=subject, body=body)
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_has_calls([
            call("c", 0),
            *[c for email in recipients for c in [*map(lambda char: call(char, 1), email), call("{ENTER}", 0)]],
            call("{TAB}", 0),
            *[call(char, 1) for char in subject],
            call("{TAB}", 0),
            *[call(char, 1) for char in body],
            call("^{ENTER}", 0),
        ])
        self.gmail.dll.AU3_WinActive.assert_has_calls([call(self.gmail.browser.window_info, "")] * (sum(len(email) for email in recipients) + len(subject) + len(body)))

        self.assertEqual(self.gmail.dll.AU3_WinActive.call_count, sum(len(email) for email in recipients) + len(subject) + len(body))
        self.assertEqual(self.gmail.dll.AU3_Send.call_count, sum(len(email) for email in recipients) + len(subject) + len(body) + len(recipients) + 4)
        self.assertEqual(mock_sleep.call_count, sum(len(email) for email in recipients) + len(subject) + len(body) + len(recipients) + 7)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    def test_compose_email_failed_to_open_compose_tab(self, mock_sleep):
        recipients = ["example1@mail.com", "example2@mail.com", "example3@mail.com"]
        subject = "Example Subject Gmail"
        body = "example body gmail"
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.return_value = 0
        
        res = self.gmail.compose_email(recipients=recipients, subject=subject, body=body)
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_has_calls([call("c", 0)])
        self.gmail.dll.AU3_WinActive.asset_not_called()

        self.assertEqual(self.gmail.dll.AU3_Send.call_count, 1)
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not send key to open compose tab")
        
    @patch("time.sleep", return_value=None)
    def test_compose_email_failed_window_inactive_while_filling_recipients(self, mock_sleep):
        recipients = ["example1@mail.com", "example2@mail.com", "example3@mail.com"]
        subject = "Example Subject Gmail"
        body = "example body gmail"
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.return_value = 1
        self.gmail.dll.AU3_WinActive.side_effect = [1] * 5 + [0]
        
        res = self.gmail.compose_email(recipients=recipients, subject=subject, body=body)
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_has_calls([
            call("c", 0),
            *[call(char, 1) for char in "examp"]
        ])
        self.gmail.dll.AU3_WinActive.assert_has_calls([call(self.gmail.browser.window_info, "")] * 6)

        self.assertEqual(self.gmail.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.gmail.dll.AU3_Send.call_count, 6)
        self.assertEqual(mock_sleep.call_count, 7)
        self.assertEqual(res, "browser window is inactive while filling email recipient")

    @patch("time.sleep", return_value=None)
    def test_compose_email_failed_to_send_letter_while_filling_recipients(self, mock_sleep):
        recipients = ["example1@mail.com", "example2@mail.com", "example3@mail.com"]
        subject = "Example Subject Gmail"
        body = "example body gmail"
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.side_effect = [1] * 6 + [0]
        self.gmail.dll.AU3_WinActive.return_value = 1
        
        res = self.gmail.compose_email(recipients=recipients, subject=subject, body=body)
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_has_calls([
            call("c", 0),
            *[call(char, 1) for char in "exampl"]
        ])
        self.gmail.dll.AU3_WinActive.assert_has_calls([call(self.gmail.browser.window_info, "")] * 6)

        self.assertEqual(self.gmail.dll.AU3_WinActive.call_count, 6)
        self.assertEqual(self.gmail.dll.AU3_Send.call_count, 7)
        self.assertEqual(mock_sleep.call_count, 7)
        self.assertEqual(res, "could not send l to recipient field")
        
    @patch("time.sleep", return_value=None)
    def test_compose_email_failed_to_send_enter_while_filling_recipients(self, mock_sleep):
        recipients = ["example1@mail.com", "example2@mail.com", "example3@mail.com"]
        subject = "Example Subject Gmail"
        body = "example body gmail"
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.side_effect = [1] * (len(recipients[0]) + 1) + [0]
        self.gmail.dll.AU3_WinActive.return_value = 1
        
        res = self.gmail.compose_email(recipients=recipients, subject=subject, body=body)
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_has_calls([
            call("c", 0),
            *[call(char, 1) for char in recipients[0]],
            call("{ENTER}", 0)
        ])
        self.gmail.dll.AU3_WinActive.assert_has_calls([call(self.gmail.browser.window_info, "")] * len(recipients[0]))

        self.assertEqual(self.gmail.dll.AU3_WinActive.call_count, len(recipients[0]))
        self.assertEqual(self.gmail.dll.AU3_Send.call_count, len(recipients[0]) + 2)
        self.assertEqual(mock_sleep.call_count, len(recipients[0]) + 3)
        self.assertEqual(res, "could not send enter key to confirm recipient")
        
    @patch("time.sleep", return_value=None)
    def test_compose_email_failed_to_send_tab_to_go_to_subject_field(self, mock_sleep):
        recipients = ["example1@mail.com", "example2@mail.com", "example3@mail.com"]
        subject = "Example Subject Gmail"
        body = "example body gmail"
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.side_effect = [1] * (sum(len(email) for email in recipients) + len(recipients) + 1) + [0]
        self.gmail.dll.AU3_WinActive.return_value = 1
        
        res = self.gmail.compose_email(recipients=recipients, subject=subject, body=body)
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_has_calls([
            call("c", 0),
            *[c for email in recipients for c in [*map(lambda char: call(char, 1), email), call("{ENTER}", 0)]],
            call("{TAB}", 0),
        ])
        self.gmail.dll.AU3_WinActive.assert_has_calls([call(self.gmail.browser.window_info, "")] * sum(len(email) for email in recipients))

        self.assertEqual(self.gmail.dll.AU3_WinActive.call_count, sum(len(email) for email in recipients))
        self.assertEqual(self.gmail.dll.AU3_Send.call_count, sum(len(email) for email in recipients) + len(recipients) + 2)
        self.assertEqual(mock_sleep.call_count, sum(len(email) for email in recipients) + len(recipients) + 3)
        self.assertEqual(res, "could not send key to go to subject field")
        
    @patch("time.sleep", return_value=None)
    def test_compose_email_failed_window_inactive_while_filling_subject(self, mock_sleep):
        recipients = ["example1@mail.com", "example2@mail.com", "example3@mail.com"]
        subject = "Example Subject Gmail"
        body = "example body gmail"
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.return_value = 1
        self.gmail.dll.AU3_WinActive.side_effect = [1] * (sum(len(email) for email in recipients) + 5) + [0]
        
        res = self.gmail.compose_email(recipients=recipients, subject=subject, body=body)
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_has_calls([
            call("c", 0),
            *[c for email in recipients for c in [*map(lambda char: call(char, 1), email), call("{ENTER}", 0)]],
            call("{TAB}", 0),
            *[call(char, 1) for char in "Examp"]
        ])
        self.gmail.dll.AU3_WinActive.assert_has_calls([call(self.gmail.browser.window_info, "")] * (sum(len(email) for email in recipients) + 6))

        self.assertEqual(self.gmail.dll.AU3_WinActive.call_count, sum(len(email) for email in recipients) + 6)
        self.assertEqual(self.gmail.dll.AU3_Send.call_count, sum(len(email) for email in recipients) + len(recipients) + 7)
        self.assertEqual(mock_sleep.call_count, sum(len(email) for email in recipients) + len(recipients) + 9)
        self.assertEqual(res, "browser window is inactive while filling email subject")
        
    @patch("time.sleep", return_value=None)
    def test_compose_email_failed_to_send_letter_while_filling_subject(self, mock_sleep):
        recipients = ["example1@mail.com", "example2@mail.com", "example3@mail.com"]
        subject = "Example Subject Gmail"
        body = "example body gmail"
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.side_effect = [1] * (sum(len(email) for email in recipients) + len(recipients) + 7) + [0]
        self.gmail.dll.AU3_WinActive.return_value = 1
        
        res = self.gmail.compose_email(recipients=recipients, subject=subject, body=body)
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_has_calls([
            call("c", 0),
            *[c for email in recipients for c in [*map(lambda char: call(char, 1), email), call("{ENTER}", 0)]],
            call("{TAB}", 0),
            *[call(char, 1) for char in "Exampl"]
        ])
        self.gmail.dll.AU3_WinActive.assert_has_calls([call(self.gmail.browser.window_info, "")] * (sum(len(email) for email in recipients) + 5))

        self.assertEqual(self.gmail.dll.AU3_WinActive.call_count, sum(len(email) for email in recipients) + 6)
        self.assertEqual(self.gmail.dll.AU3_Send.call_count, sum(len(email) for email in recipients) + len(recipients) + 8)
        self.assertEqual(mock_sleep.call_count, sum(len(email) for email in recipients) + len(recipients) + 9)
        self.assertEqual(res, "could not send l to subject field")
        
    @patch("time.sleep", return_value=None)
    def test_compose_email_failed_to_send_tab_to_go_to_body_field(self, mock_sleep):
        recipients = ["example1@mail.com", "example2@mail.com", "example3@mail.com"]
        subject = "Example Subject Gmail"
        body = "example body gmail"
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.side_effect = [1] * (sum(len(email) for email in recipients) + len(recipients) + len(subject) + 2) + [0]
        self.gmail.dll.AU3_WinActive.return_value = 1
        
        res = self.gmail.compose_email(recipients=recipients, subject=subject, body=body)
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_has_calls([
            call("c", 0),
            *[c for email in recipients for c in [*map(lambda char: call(char, 1), email), call("{ENTER}", 0)]],
            call("{TAB}", 0),
            *[call(char, 1) for char in subject],
            call("{TAB}", 0),
        ])
        self.gmail.dll.AU3_WinActive.assert_has_calls([call(self.gmail.browser.window_info, "")] * (sum(len(email) for email in recipients) + len(subject)))

        self.assertEqual(self.gmail.dll.AU3_WinActive.call_count, sum(len(email) for email in recipients) + len(subject))
        self.assertEqual(self.gmail.dll.AU3_Send.call_count, sum(len(email) for email in recipients) + len(recipients) + len(subject) + 3)
        self.assertEqual(mock_sleep.call_count, sum(len(email) for email in recipients) + len(recipients) + len(subject) + 5)
        self.assertEqual(res, "could not send key to go to body field")
        
    @patch("time.sleep", return_value=None)
    def test_compose_email_failed_window_inactive_while_filling_body(self, mock_sleep):
        recipients = ["example1@mail.com", "example2@mail.com", "example3@mail.com"]
        subject = "Example Subject Gmail"
        body = "example body gmail"
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.return_value = 1
        self.gmail.dll.AU3_WinActive.side_effect = [1] * (sum(len(email) for email in recipients) + len(subject) + 5) + [0]
        
        res = self.gmail.compose_email(recipients=recipients, subject=subject, body=body)
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_has_calls([
            call("c", 0),
            *[c for email in recipients for c in [*map(lambda char: call(char, 1), email), call("{ENTER}", 0)]],
            call("{TAB}", 0),
            *[call(char, 1) for char in subject],
            call("{TAB}", 0),
            *[call(char, 1) for char in "examp"],
        ])
        self.gmail.dll.AU3_WinActive.assert_has_calls([call(self.gmail.browser.window_info, "")] * (sum(len(email) for email in recipients) + len(subject) + 6))

        self.assertEqual(self.gmail.dll.AU3_WinActive.call_count, sum(len(email) for email in recipients) + len(subject) + 6)
        self.assertEqual(self.gmail.dll.AU3_Send.call_count, sum(len(email) for email in recipients) + len(recipients) + len(subject) + 8)
        self.assertEqual(mock_sleep.call_count, sum(len(email) for email in recipients) + len(recipients) + len(subject) + 11)
        self.assertEqual(res, "browser window is inactive while filling email body")
        
    @patch("time.sleep", return_value=None)
    def test_compose_email_failed_to_send_letter_while_filling_body(self, mock_sleep):
        recipients = ["example1@mail.com", "example2@mail.com", "example3@mail.com"]
        subject = "Example Subject Gmail"
        body = "example body gmail"
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.side_effect = [1] * (sum(len(email) for email in recipients) + len(recipients) + len(subject) + 8) + [0]
        self.gmail.dll.AU3_WinActive.return_value = 1
        
        res = self.gmail.compose_email(recipients=recipients, subject=subject, body=body)
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_has_calls([
            call("c", 0),
            *[c for email in recipients for c in [*map(lambda char: call(char, 1), email), call("{ENTER}", 0)]],
            call("{TAB}", 0),
            *[call(char, 1) for char in subject],
            call("{TAB}", 0),
            *[call(char, 1) for char in "exampl"],
        ])
        self.gmail.dll.AU3_WinActive.assert_has_calls([call(self.gmail.browser.window_info, "")] * (sum(len(email) for email in recipients) + len(subject) + 6))

        self.assertEqual(self.gmail.dll.AU3_WinActive.call_count, sum(len(email) for email in recipients) + len(subject) + 6)
        self.assertEqual(self.gmail.dll.AU3_Send.call_count, sum(len(email) for email in recipients) + len(recipients) + len(subject) + 9)
        self.assertEqual(mock_sleep.call_count, sum(len(email) for email in recipients) + len(recipients) + len(subject) + 11)
        self.assertEqual(res, "could not send l to body field")
        
    @patch("time.sleep", return_value=None)
    def test_compose_email_failed_to_send_letter_while_filling_body(self, mock_sleep):
        recipients = ["example1@mail.com", "example2@mail.com", "example3@mail.com"]
        subject = "Example Subject Gmail"
        body = "example body gmail"
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.side_effect = [1] * (sum(len(email) for email in recipients) + len(recipients) + len(subject) + len(body) + 3) + [0]
        self.gmail.dll.AU3_WinActive.return_value = 1
        
        res = self.gmail.compose_email(recipients=recipients, subject=subject, body=body)
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_has_calls([
            call("c", 0),
            *[c for email in recipients for c in [*map(lambda char: call(char, 1), email), call("{ENTER}", 0)]],
            call("{TAB}", 0),
            *[call(char, 1) for char in subject],
            call("{TAB}", 0),
            *[call(char, 1) for char in body],
            call("^{ENTER}", 0),
        ])
        self.gmail.dll.AU3_WinActive.assert_has_calls([call(self.gmail.browser.window_info, "")] * (sum(len(email) for email in recipients) + len(subject) + len(body)))

        self.assertEqual(self.gmail.dll.AU3_WinActive.call_count, sum(len(email) for email in recipients) + len(subject) + len(body))
        self.assertEqual(self.gmail.dll.AU3_Send.call_count, sum(len(email) for email in recipients) + len(subject) + len(body) + len(recipients) + 4)
        self.assertEqual(mock_sleep.call_count, sum(len(email) for email in recipients) + len(subject) + len(body) + len(recipients) + 7)
        self.assertEqual(res, "could not send keys to send email")
        
    @patch("time.sleep", return_value=None)
    def test_go_to_inbox_success(self, mock_sleep):
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.return_value = 1
        
        res = self.gmail.go_to_inbox()
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_called_once_with("gi", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    def test_go_to_inbox_failed_to_send_keys(self, mock_sleep):
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.return_value = 0
        
        res = self.gmail.go_to_inbox()
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_called_once_with("gi", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not send keys to go to inbox")
        
    @patch("time.sleep", return_value=None)
    def test_go_to_drafts_success(self, mock_sleep):
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.return_value = 1
        
        res = self.gmail.go_to_drafts()
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_called_once_with("gd", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    def test_go_to_drafts_failed_to_send_keys(self, mock_sleep):
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.return_value = 0
        
        res = self.gmail.go_to_drafts()
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_called_once_with("gd", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not send keys to go to drafts")
        
    @patch("time.sleep", return_value=None)
    def test_go_to_sent_messages_success(self, mock_sleep):
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.return_value = 1
        
        res = self.gmail.go_to_sent_messages()
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_called_once_with("gt", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    def test_go_to_sent_messages_failed_to_send_keys(self, mock_sleep):
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.return_value = 0
        
        res = self.gmail.go_to_sent_messages()
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_called_once_with("gt", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not send keys to go to sent messages")
        
    @patch("time.sleep", return_value=None)
    def test_select_all_success(self, mock_sleep):
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.return_value = 1
        
        res = self.gmail.select_all()
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_called_once_with("*a", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    def test_select_all_failed_to_send_keys(self, mock_sleep):
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.return_value = 0
        
        res = self.gmail.select_all()
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_called_once_with("*a", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not send keys to select all email")
        
    @patch("time.sleep", return_value=None)
    def test_deselect_all_success(self, mock_sleep):
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.return_value = 1
        
        res = self.gmail.deselect_all()
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_called_once_with("*n", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    def test_deselect_all_failed_to_send_keys(self, mock_sleep):
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.return_value = 0
        
        res = self.gmail.deselect_all()
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_called_once_with("*n", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not send keys to deselect all email")
        
    @patch("time.sleep", return_value=None)
    def test_select_unread_success(self, mock_sleep):
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.return_value = 1
        
        res = self.gmail.select_unread()
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_called_once_with("*u", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    def test_select_unread_failed_to_send_keys(self, mock_sleep):
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.return_value = 0
        
        res = self.gmail.select_unread()
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_called_once_with("*u", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not send keys to select unread email")
        
    @patch("time.sleep", return_value=None)
    def test_mark_as_unread_success(self, mock_sleep):
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.return_value = 1
        
        res = self.gmail.mark_as_unread()
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_called_once_with("+u", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    def test_mark_as_unread_failed_to_send_keys(self, mock_sleep):
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.return_value = 0
        
        res = self.gmail.mark_as_unread()
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_called_once_with("+u", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not send keys to mark selected email as unread")
        
    @patch("time.sleep", return_value=None)
    def test_mark_as_read_success(self, mock_sleep):
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.return_value = 1
        
        res = self.gmail.mark_as_read()
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_called_once_with("+i", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    def test_mark_as_read_failed_to_send_keys(self, mock_sleep):
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.return_value = 0
        
        res = self.gmail.mark_as_read()
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_called_once_with("+i", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not send keys to mark selected email as read")
        
    @patch("time.sleep", return_value=None)
    def test_archive_email_success(self, mock_sleep):
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.return_value = 1
        
        res = self.gmail.archive_email()
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_called_once_with("e", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    def test_archive_email_failed_to_send_key(self, mock_sleep):
        self.gmail.browser.check_existing_window.return_value = None
        self.gmail.dll.AU3_Send.return_value = 0
        
        res = self.gmail.archive_email()
        self.gmail.browser.check_existing_window.assert_called_once()
        self.gmail.dll.AU3_Send.assert_called_once_with("e", 0)
        
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not send key to archive email")