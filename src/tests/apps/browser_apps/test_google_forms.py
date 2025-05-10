import unittest
from unittest.mock import patch, call, MagicMock
from activities.apps.browser_apps.google_forms import GoogleForms

class TestGoogleForms(unittest.TestCase):
    def setUp(self):
        self.gforms = GoogleForms(browser=MagicMock())
        self.gforms.dll = MagicMock()
    
    @patch("time.sleep", return_value=None)    
    def test_fill_form_success_with_create_window(self, mock_sleep):
        url = "https://forms.gle/abcdefghijklmnopq"
        answers = ["short answer", "lorem ipsum", "paragraph answer"]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.dll.AU3_WinActive.return_value = 1
        self.gforms.dll.AU3_Send.return_value = 1
        
        res = self.gforms.fill_form(url=url, answers=answers)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with(url)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(answers) + sum(len(answer) for answer in answers)))
        self.gforms.dll.AU3_Send.assert_has_calls([
            call("{TAB}{TAB}{TAB}", 0),
            *[
                c
                for i, answer in enumerate(answers)
                for c in (
                    ([call("{TAB}", 0)] if i > 0 else [])
                    + list(map(lambda char: call(char, 1), answer))
                )
            ],
            call("{TAB}{ENTER}", 0)
        ])
        self.assertEqual(mock_sleep.call_count, len(answers) + sum(len(answer) for answer in answers) + 2)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    def test_fill_form_success_with_create_tab(self, mock_sleep):
        url = "https://forms.gle/abcdefghijklmnopq"
        answers = ["short answer", "lorem ipsum", "paragraph answer"]
        self.gforms.browser.check_existing_window.return_value = None
        self.gforms.browser.create_tab.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.dll.AU3_WinActive.return_value = 1
        self.gforms.dll.AU3_Send.return_value = 1
        
        res = self.gforms.fill_form(url=url, answers=answers)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_not_called()
        self.gforms.browser.create_tab.assert_called_once()
        self.gforms.browser.browse.assert_called_once_with(url)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(answers) + sum(len(answer) for answer in answers)))
        self.gforms.dll.AU3_Send.assert_has_calls([
            call("{TAB}{TAB}{TAB}", 0),
            *[
                c
                for i, answer in enumerate(answers)
                for c in (
                    ([call("{TAB}", 0)] if i > 0 else [])
                    + list(map(lambda char: call(char, 1), answer))
                )
            ],
            call("{TAB}{ENTER}", 0)
        ])
        self.assertEqual(mock_sleep.call_count, len(answers) + sum(len(answer) for answer in answers) + 2)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    def test_fill_form_invalid_url(self, mock_sleep):
        url = "https://not/gform/url"
        answers = ["short answer", "lorem ipsum", "paragraph answer"]
    
        res = self.gforms.fill_form(url=url, answers=answers)

        self.gforms.browser.check_existing_window.assert_not_called()
        self.gforms.browser.create_window.assert_not_called()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_not_called()
        self.gforms.dll.AU3_WinActive.assert_not_called()
        self.gforms.dll.AU3_Send.assert_not_called()
        self.assertEqual(mock_sleep.call_count, 0)
        self.assertEqual(res, "Invalid Google Forms URL")
        
    @patch("time.sleep", return_value=None)
    def test_fill_form_failed_window_inactive_while_moving_to_questions(self, mock_sleep):
        url = "https://forms.gle/abcdefghijklmnopq"
        answers = ["short answer", "lorem ipsum", "paragraph answer"]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.dll.AU3_WinActive.side_effect = [1] * (len(answers[0]) + 1) + [0]
        self.gforms.dll.AU3_Send.return_value = 1
        
        res = self.gforms.fill_form(url=url, answers=answers)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with(url)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * 6)
        self.gforms.dll.AU3_Send.assert_has_calls([
            call("{TAB}{TAB}{TAB}", 0),
            *[call(char, 1) for char in answers[0]],
        ])
        self.assertEqual(mock_sleep.call_count, len(answers[0]) + 2)
        self.assertEqual(res, "browser window is inactive while moving to Google Forms questions")
        
    @patch("time.sleep", return_value=None)
    def test_fill_form_failed_to_send_key_to_move_to_first_question(self, mock_sleep):
        url = "https://forms.gle/abcdefghijklmnopq"
        answers = ["short answer", "lorem ipsum", "paragraph answer"]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.dll.AU3_WinActive.return_value = 1
        self.gforms.dll.AU3_Send.return_value = 0
        
        res = self.gforms.fill_form(url=url, answers=answers)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with(url)
        self.gforms.dll.AU3_WinActive.assert_called_once_with(self.gforms.browser.window_info, "")
        self.gforms.dll.AU3_Send.assert_has_calls([
            call("{TAB}{TAB}{TAB}", 0),
        ])
        self.assertEqual(mock_sleep.call_count, 1)
        self.assertEqual(res, "could not send tab key to move to Google Forms question")
        
    @patch("time.sleep", return_value=None)
    def test_fill_form_failed_to_send_key_to_move_to_next_question(self, mock_sleep):
        url = "https://forms.gle/abcdefghijklmnopq"
        answers = ["short answer", "lorem ipsum", "paragraph answer"]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.dll.AU3_WinActive.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (len(answers[0]) + 1) + [0]
        
        res = self.gforms.fill_form(url=url, answers=answers)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with(url)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * 6)
        self.gforms.dll.AU3_Send.assert_has_calls([
            call("{TAB}{TAB}{TAB}", 0),
            *[call(char, 1) for char in answers[0]],
            call("{TAB}", 0),
        ])
        self.assertEqual(mock_sleep.call_count, len(answers[0]) + 2)
        self.assertEqual(res, "could not send tab key to move to next question")
        
    @patch("time.sleep", return_value=None)
    def test_fill_form_failed_window_inactive_while_filling_answers(self, mock_sleep):
        url = "https://forms.gle/abcdefghijklmnopq"
        answers = ["short answer", "lorem ipsum", "paragraph answer"]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.dll.AU3_WinActive.side_effect = [1, 0]
        self.gforms.dll.AU3_Send.return_value = 1
        
        res = self.gforms.fill_form(url=url, answers=answers)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with(url)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * 2)
        self.gforms.dll.AU3_Send.assert_has_calls([
            call("{TAB}{TAB}{TAB}", 0),
        ])
        self.assertEqual(mock_sleep.call_count, 2)
        self.assertEqual(res, "browser window is inactive while filling Google Forms answers")
        
    @patch("time.sleep", return_value=None)
    def test_fill_form_failed_to_send_letter_while_filling_answers(self, mock_sleep):
        url = "https://forms.gle/abcdefghijklmnopq"
        answers = ["short answer", "lorem ipsum", "paragraph answer"]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.dll.AU3_WinActive.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (len(answers[0]) + 6) + [0]
        
        res = self.gforms.fill_form(url=url, answers=answers)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with(url)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(answers[0]) + 6))
        self.gforms.dll.AU3_Send.assert_has_calls([
            call("{TAB}{TAB}{TAB}", 0),
            *[call(char, 1) for char in answers[0]],
            call("{TAB}", 0),
            *[call(char, 1) for char in "lorem"],
        ])
        self.assertEqual(mock_sleep.call_count, len(answers[0]) + 7)
        self.assertEqual(res, "could not send m to Google Forms")

    @patch("time.sleep", return_value=None)
    def test_fill_form_failed_to_send_keys_to_sumbit(self, mock_sleep):
        url = "https://forms.gle/abcdefghijklmnopq"
        answers = ["short answer", "lorem ipsum", "paragraph answer"]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.dll.AU3_WinActive.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (sum(len(answer) for answer in answers) + len(answers)) + [0]
        
        res = self.gforms.fill_form(url=url, answers=answers)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with(url)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (sum(len(answer) for answer in answers) + len(answers)))
        self.gforms.dll.AU3_Send.assert_has_calls([
            call("{TAB}{TAB}{TAB}", 0),
            *[
                c
                for i, answer in enumerate(answers)
                for c in (
                    ([call("{TAB}", 0)] if i > 0 else [])
                    + list(map(lambda char: call(char, 1), answer))
                )
            ],
            call("{TAB}{ENTER}", 0)
        ])
        self.assertEqual(mock_sleep.call_count, (sum(len(answer) for answer in answers) + len(answers) + 2))
        self.assertEqual(res, "could not send keys to submit Google Forms")
    
    @patch("time.sleep", return_value=None)
    def test_create_form_success_with_create_window(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.return_value = 1
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_has_calls([call(), call()])
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in title] +
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in description]
        )
        
        if len(questions) > 1:
            expected_send_calls += [call("+{TAB}", 0)] * 8 + [call("{ENTER}", 0)]
        if len(questions) > 2:
            expected_send_calls += ([call("+{TAB}", 0)] * 10 + [call("{ENTER}", 0)]) * (len(questions) - 2)

        for i, (question_type, question_title, is_required) in enumerate(questions):
            expected_send_calls += (
                [call(char, 1) for char in question_title] +
                [call("{TAB}", 0)] * 3 +
                [call("{ENTER}", 0)] +
                [call("{UP}{UP}", 0) if question_type == "Short answer" else call("{UP}", 0)] +
                [call("{ENTER}", 0)] +
                [call("{TAB}", 0)] * 3 +
                ([call("{SPACE}", 0)] if is_required else []) +
                (([call("{TAB}", 0)] * 2) if i < len(questions) - 1 else [])
            )

        expected_send_calls += (
            [call("+{TAB}", 0)] * (20 + (len(questions) - 1) * 2) +
            [call("{ENTER}", 0), call("{TAB}{TAB}", 0), call("{ENTER}", 0)] +
            [call("+{TAB}", 0)] * 2 +
            [call("{SPACE}", 0)] +
            [call("{TAB}", 0)] +
            [call("{SPACE}", 0)] +
            [call("{TAB}", 0)] * 2 +
            [call("{SPACE}", 0)]
        )

        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + len(title) + len(description) + len(questions) - 1 + sum(len(question[1]) for question in questions)))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + len(title) +
            3 + 1 + len(description) +
            1 + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + 1 +
            sum([
                len(q_title) +
                3 + 1 + 1 + 1 + 3 +
                (2 if i < len(questions) - 1 else (20 + (len(questions) - 1) * 2 + 1 + 2 + 1))
                for i, (_, q_title, _) in enumerate(questions)
            ]) +
            2 + 1 + 1 + 1 + 2 + 1 + 1 + 2
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    def test_create_form_success_with_create_tab(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = None
        self.gforms.browser.create_tab.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.return_value = 1
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_not_called()
        self.gforms.browser.create_tab.assert_called_once()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_has_calls([call(), call()])
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in title] +
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in description]
        )
        
        if len(questions) > 1:
            expected_send_calls += [call("+{TAB}", 0)] * 8 + [call("{ENTER}", 0)]
        if len(questions) > 2:
            expected_send_calls += ([call("+{TAB}", 0)] * 10 + [call("{ENTER}", 0)]) * (len(questions) - 2)

        for i, (question_type, question_title, is_required) in enumerate(questions):
            expected_send_calls += (
                [call(char, 1) for char in question_title] +
                [call("{TAB}", 0)] * 3 +
                [call("{ENTER}", 0)] +
                [call("{UP}{UP}", 0) if question_type == "Short answer" else call("{UP}", 0)] +
                [call("{ENTER}", 0)] +
                [call("{TAB}", 0)] * 3 +
                ([call("{SPACE}", 0)] if is_required else []) +
                (([call("{TAB}", 0)] * 2) if i < len(questions) - 1 else [])
            )

        expected_send_calls += (
            [call("+{TAB}", 0)] * (20 + (len(questions) - 1) * 2) +
            [call("{ENTER}", 0), call("{TAB}{TAB}", 0), call("{ENTER}", 0)] +
            [call("+{TAB}", 0)] * 2 +
            [call("{SPACE}", 0)] +
            [call("{TAB}", 0)] +
            [call("{SPACE}", 0)] +
            [call("{TAB}", 0)] * 2 +
            [call("{SPACE}", 0)]
        )

        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + len(title) + len(description) + len(questions) - 1 + sum(len(question[1]) for question in questions)))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + len(title) +
            3 + 1 + len(description) +
            1 + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + 1 +
            sum([
                len(q_title) +
                3 + 1 + 1 + 1 + 3 +
                (2 if i < len(questions) - 1 else (20 + (len(questions) - 1) * 2 + 1 + 2 + 1))
                for i, (_, q_title, _) in enumerate(questions)
            ]) +
            2 + 1 + 1 + 1 + 2 + 1 + 1 + 2
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertIsNone(res)
        
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_click_on_top_left(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 0

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        self.gforms.dll.AU3_Send.assert_not_called()
        self.gforms.dll.AU3_WinActive.assert_not_called()
        
        self.assertEqual(mock_sleep.call_count, 2)
        self.assertEqual(res, "failed to click on top left")
        
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_send_key_to_move_to_form_name(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1, 0]

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2
        )
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_not_called()
        expected_sleep_calls = (
            5
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "could not send tab key to move to Google Forms name")
        
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_window_inactive_while_filling_form_name(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.return_value = 1
        self.gforms.dll.AU3_WinActive.side_effect = [1] * 5 + [0]

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in "Examp"]
        )
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * 6)
        
        expected_sleep_calls = (
            5 + 1 + 5
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "browser window is inactive while filling Google Forms name")
        
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_send_letter_while_filling_form_name(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * 7 + [0]
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in "Exampl"]
        )
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * 6)
        
        expected_sleep_calls = (
            5 + 1 + 5
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "could not send l to Google Forms")
        
    
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_send_key_to_move_to_form_title(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (len(name) + 5) + [0]
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 4
        )

        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * len(name))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            5
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "could not send tab key to move to Google Forms title")
        
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_clear_form_title(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (len(name) + 20) + [0]
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)]
        )
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * len(name))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "could not send keys to delete template Google Forms title")
        
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_window_inactive_while_filling_form_title(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.return_value = 1
        self.gforms.dll.AU3_WinActive.side_effect = [1] * (len(name) + 5) + [0]

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in "Examp"]
        )
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + 6))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + 5
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "browser window is inactive while filling Google Forms title")
        
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_send_letter_while_filling_form_title(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (len(name) + 25) + [0]
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in "Examp"]
        )
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + 5))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + 4
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "could not send p to Google Forms")
        
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_send_key_to_move_to_form_description(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (len(name) + len(title) + 22) + [0]
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in title] +
            [call("{TAB}", 0)] * 2
        )
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + len(title)))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + len(title) +
            3
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res,"could not send tab key to move to Google Forms description")
        
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_window_inactive_while_filling_form_description(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.return_value = 1
        self.gforms.dll.AU3_WinActive.side_effect = [1] * (len(name) + len(title) + 5) + [0]

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in title] +
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in "Examp"]
        )
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + len(title) + 6))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + len(title) +
            3 + 1 + 5
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "browser window is inactive while filling Google Forms description")
        
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_send_letter_while_filling_form_description(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (len(name) + len(title) + 27) + [0]
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in title] +
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in "Examp"]
        )
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + len(title) + 5))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + len(title) +
            3 + 1 + 4
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "could not send p to Google Forms")
        
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_window_inactive_while_adding_question_boxes(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.return_value = 1
        self.gforms.dll.AU3_WinActive.side_effect = [1] * (len(name) + len(title) + len(description) + 1) + [0]

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in title] +
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in description]
        )
        expected_send_calls += [call("+{TAB}", 0)] * 8 + [call("{ENTER}", 0)]
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + len(title) + len(description) + 2))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + len(title) +
            3 + 1 + len(description) +
            1 + 8
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "browser window is inactive while adding Google Forms question box")
        
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_send_keys_to_move_to_create_first_question_box_button(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (len(name) + len(title) + len(description) + 27) + [0]
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in title] +
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in description]
        )
        expected_send_calls += [call("+{TAB}", 0)] * 5
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + len(title) + len(description) + 1))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + len(title) +
            3 + 1 + len(description) +
            1 + 5
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "could not send keys to move to create question box button")
         
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_send_keys_to_move_to_create_next_question_box_button(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (len(name) + len(title) + len(description) + 36) + [0]
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in title] +
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in description]
        )
        expected_send_calls += [call("+{TAB}", 0)] * 8 + [call("{ENTER}", 0)]
        expected_send_calls += [call("+{TAB}", 0)] * 5
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + len(title) + len(description) + 2))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + len(title) +
            3 + 1 + len(description) +
            1 + 8 + 5
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "could not send keys to move to create question box button")
        
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_send_key_to_press_question_box_button(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (len(name) + len(title) + len(description) + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + 24) + [0]
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in title] +
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in description]
        )
        if len(questions) > 1:
            expected_send_calls += [call("+{TAB}", 0)] * 8 + [call("{ENTER}", 0)]
        if len(questions) > 2:
            expected_send_calls += ([call("+{TAB}", 0)] * 10 + [call("{ENTER}", 0)]) * (len(questions) - 2)
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + len(title) + len(description) + len(questions) - 1))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + len(title) +
            3 + 1 + len(description) +
            1 + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0)
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "could not press question box button")
        
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_window_inactive_while_filling_question_title(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.return_value = 1
        self.gforms.dll.AU3_WinActive.side_effect = [1] * (len(name) + len(title) + len(description) + len(questions) - 1 + len(questions[0][1])) + [0]

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in title] +
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in description]
        )  
        if len(questions) > 1:
            expected_send_calls += [call("+{TAB}", 0)] * 8 + [call("{ENTER}", 0)]
        if len(questions) > 2:
            expected_send_calls += ([call("+{TAB}", 0)] * 10 + [call("{ENTER}", 0)]) * (len(questions) - 2)
        expected_send_calls += [call(char, 1) for char in questions[0][1]]
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + len(title) + len(description) + len(questions) - 1 + len(questions[0][1])))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + len(title) +
            3 + 1 + len(description) +
            1 + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + 1 + len(questions[0][1]) + 3 + 1 + 1 + 1 + 3 + 2
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "browser window is inactive while filling question title")
        
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_send_letter_while_filling_question_title(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (len(name) + len(title) + len(description) + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + len(questions[0][1]) + 24) + [0]
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in title] +
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in description]
        )  
        if len(questions) > 1:
            expected_send_calls += [call("+{TAB}", 0)] * 8 + [call("{ENTER}", 0)]
        if len(questions) > 2:
            expected_send_calls += ([call("+{TAB}", 0)] * 10 + [call("{ENTER}", 0)]) * (len(questions) - 2)
        expected_send_calls += [call(char, 1) for char in questions[0][1]]
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + len(title) + len(description) + len(questions) - 1 + len(questions[0][1])))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + len(title) +
            3 + 1 + len(description) +
            1 + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + 1 + len(questions[0][1]) - 1
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, f"could not send 1 to Google Forms")
        
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_send_key_to_move_to_question_type_dropdown(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (len(name) + len(title) + len(description) + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + len(questions[0][1]) + 27) + [0]
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in title] +
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in description]
        )   
        if len(questions) > 1:
            expected_send_calls += [call("+{TAB}", 0)] * 8 + [call("{ENTER}", 0)]
        if len(questions) > 2:
            expected_send_calls += ([call("+{TAB}", 0)] * 10 + [call("{ENTER}", 0)]) * (len(questions) - 2)
        expected_send_calls += (
            [call(char, 1) for char in questions[0][1]] +
            [call("{TAB}", 0)]
        )
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + len(title) + len(description) + len(questions) - 1 + len(questions[0][1])))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + len(title) +
            3 + 1 + len(description) +
            1 + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + 1 + len(questions[0][1]) + 3
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "could not send tab key to move to question type dropdown")
        
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_send_key_to_make_question_type_dropdown_appear(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (len(name) + len(title) + len(description) + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + len(questions[0][1]) + 28) + [0]
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in title] +
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in description]
        )   
        if len(questions) > 1:
            expected_send_calls += [call("+{TAB}", 0)] * 8 + [call("{ENTER}", 0)]
        if len(questions) > 2:
            expected_send_calls += ([call("+{TAB}", 0)] * 10 + [call("{ENTER}", 0)]) * (len(questions) - 2)
        expected_send_calls += (
            [call(char, 1) for char in questions[0][1]] +
            [call("{TAB}", 0)] * 3 +
            [call("{ENTER}", 0)]
        )
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + len(title) + len(description) + len(questions) - 1 + len(questions[0][1])))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + len(title) +
            3 + 1 + len(description) +
            1 + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + 1 + len(questions[0][1]) + 3 + 1
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "could not send the Enter key to make the question type dropdown appear")
        
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_send_key_to_move_to_question_type_short_answer(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (len(name) + len(title) + len(description) + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + len(questions[0][1]) + 29) + [0]
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in title] +
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in description]
        )   
        if len(questions) > 1:
            expected_send_calls += [call("+{TAB}", 0)] * 8 + [call("{ENTER}", 0)]
        if len(questions) > 2:
            expected_send_calls += ([call("+{TAB}", 0)] * 10 + [call("{ENTER}", 0)]) * (len(questions) - 2)
        expected_send_calls += (
            [call(char, 1) for char in questions[0][1]] +
            [call("{TAB}", 0)] * 3 +
            [call("{ENTER}", 0)] +
            [call("{UP}{UP}", 0)]
        )
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + len(title) + len(description) + len(questions) - 1 + len(questions[0][1])))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + len(title) +
            3 + 1 + len(description) +
            1 + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + 1 + len(questions[0][1]) + 3 + 1 + 1
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "could not send up key to move question type dropdown selection")
        
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_send_key_to_move_to_question_type_paragraph(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Paragraph", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (len(name) + len(title) + len(description) + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + len(questions[0][1]) + 29) + [0]
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in title] +
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in description]
        )   
        if len(questions) > 1:
            expected_send_calls += [call("+{TAB}", 0)] * 8 + [call("{ENTER}", 0)]
        if len(questions) > 2:
            expected_send_calls += ([call("+{TAB}", 0)] * 10 + [call("{ENTER}", 0)]) * (len(questions) - 2)
        expected_send_calls += (
            [call(char, 1) for char in questions[0][1]] +
            [call("{TAB}", 0)] * 3 +
            [call("{ENTER}", 0)] +
            [call("{UP}", 0)]
        )
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + len(title) + len(description) + len(questions) - 1 + len(questions[0][1])))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + len(title) +
            3 + 1 + len(description) +
            1 + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + 1 + len(questions[0][1]) + 3 + 1 + 1
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "could not send up key to move question type dropdown selection")
        
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_send_key_to_confirm_question_type_chosen(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (len(name) + len(title) + len(description) + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + len(questions[0][1]) + 30) + [0]
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in title] +
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in description]
        )   
        if len(questions) > 1:
            expected_send_calls += [call("+{TAB}", 0)] * 8 + [call("{ENTER}", 0)]
        if len(questions) > 2:
            expected_send_calls += ([call("+{TAB}", 0)] * 10 + [call("{ENTER}", 0)]) * (len(questions) - 2)
        expected_send_calls += (
            [call(char, 1) for char in questions[0][1]] +
            [call("{TAB}", 0)] * 3 +
            [call("{ENTER}", 0)] +
            [call("{UP}{UP}", 0)] +
            [call("{ENTER}", 0)]
        )
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + len(title) + len(description) + len(questions) - 1 + len(questions[0][1])))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + len(title) +
            3 + 1 + len(description) +
            1 + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + 1 + len(questions[0][1]) + 3 + 1 + 1 + 1
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "could not send enter key to choose question type from dropdown")
    
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_send_keys_to_move_to_required_toggle(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (len(name) + len(title) + len(description) + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + len(questions[0][1]) + 33) + [0]
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in title] +
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in description]
        )   
        if len(questions) > 1:
            expected_send_calls += [call("+{TAB}", 0)] * 8 + [call("{ENTER}", 0)]
        if len(questions) > 2:
            expected_send_calls += ([call("+{TAB}", 0)] * 10 + [call("{ENTER}", 0)]) * (len(questions) - 2)
        expected_send_calls += (
            [call(char, 1) for char in questions[0][1]] +
            [call("{TAB}", 0)] * 3 +
            [call("{ENTER}", 0)] +
            [call("{UP}{UP}", 0)] +
            [call("{ENTER}", 0)] +
            [call("{TAB}", 0)] * 3
        )
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + len(title) + len(description) + len(questions) - 1 + len(questions[0][1])))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + len(title) +
            3 + 1 + len(description) +
            1 + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + 1 + len(questions[0][1]) + 3 + 1 + 1 + 1 + 3
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "could not move to required toggle")
        
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_send_key_to_toggle_required(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (len(name) + len(title) + len(description) + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + len(questions[0][1]) + 34) + [0]
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in title] +
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in description]
        )   
        if len(questions) > 1:
            expected_send_calls += [call("+{TAB}", 0)] * 8 + [call("{ENTER}", 0)]
        if len(questions) > 2:
            expected_send_calls += ([call("+{TAB}", 0)] * 10 + [call("{ENTER}", 0)]) * (len(questions) - 2)
        expected_send_calls += (
            [call(char, 1) for char in questions[0][1]] +
            [call("{TAB}", 0)] * 3 +
            [call("{ENTER}", 0)] +
            [call("{UP}{UP}", 0)] +
            [call("{ENTER}", 0)] +
            [call("{TAB}", 0)] * 3 +
            [call("{SPACE}", 0)]
        )
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + len(title) + len(description) + len(questions) - 1 + len(questions[0][1])))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + len(title) +
            3 + 1 + len(description) +
            1 + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + 1 + len(questions[0][1]) + 3 + 1 + 1 + 1 + 3
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "could not toggle required")

    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_send_key_to_move_to_next_question_box(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (len(name) + len(title) + len(description) + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + len(questions[0][1]) + 36) + [0]
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in title] +
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in description]
        )   
        if len(questions) > 1:
            expected_send_calls += [call("+{TAB}", 0)] * 8 + [call("{ENTER}", 0)]
        if len(questions) > 2:
            expected_send_calls += ([call("+{TAB}", 0)] * 10 + [call("{ENTER}", 0)]) * (len(questions) - 2)
        expected_send_calls += (
            [call(char, 1) for char in questions[0][1]] +
            [call("{TAB}", 0)] * 3 +
            [call("{ENTER}", 0)] +
            [call("{UP}{UP}", 0)] +
            [call("{ENTER}", 0)] +
            [call("{TAB}", 0)] * 3 +
            [call("{SPACE}", 0)] +
            [call("{TAB}", 0)] * 2
        )
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + len(title) + len(description) + len(questions) - 1 + len(questions[0][1])))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + len(title) +
            3 + 1 + len(description) +
            1 + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + 1 + len(questions[0][1]) + 3 + 1 + 1 + 1 + 3 + 2
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "could not send tab key to move to next question box")
        
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_send_key_to_move_to_publish_button(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (len(name) + len(title) + len(description) + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + len(questions[0][1]) + 
        sum([
            len(q_title) +
            3 + 1 + 1 + 1 + 3 +
            (2 if i < len(questions) - 1 else (20 + (len(questions) - 1) * 2 + 1 + 2 + 1))
            for i, (_, q_title, _) in enumerate(questions)
        ]) + 12) + [0]
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in title] +
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in description]
        )  
        if len(questions) > 1:
            expected_send_calls += [call("+{TAB}", 0)] * 8 + [call("{ENTER}", 0)]
        if len(questions) > 2:
            expected_send_calls += ([call("+{TAB}", 0)] * 10 + [call("{ENTER}", 0)]) * (len(questions) - 2)
        for i, (question_type, question_title, is_required) in enumerate(questions):
            expected_send_calls += (
                [call(char, 1) for char in question_title] +
                [call("{TAB}", 0)] * 3 +
                [call("{ENTER}", 0)] +
                [call("{UP}{UP}", 0) if question_type == "Short answer" else call("{UP}", 0)] +
                [call("{ENTER}", 0)] +
                [call("{TAB}", 0)] * 3 +
                ([call("{SPACE}", 0)] if is_required else []) +
                (([call("{TAB}", 0)] * 2) if i < len(questions) - 1 else [])
            )
        expected_send_calls += (
            [call("+{TAB}", 0)] * (20 + (len(questions) - 1) * 2)
        )
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + len(title) + len(description) + len(questions) - 1 + sum(len(question[1]) for question in questions)))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + len(title) +
            3 + 1 + len(description) +
            1 + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + 1 +
            sum([
                len(q_title) +
                3 + 1 + 1 + 1 + 3 +
                (2 if i < len(questions) - 1 else (20 + (len(questions) - 1) * 2))
                for i, (_, q_title, _) in enumerate(questions)
            ])
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "could not send keys to move to publish button")
        
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_send_key_to_press_publish_button(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (len(name) + len(title) + len(description) + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + len(questions[0][1]) + 
        sum([
            len(q_title) +
            3 + 1 + 1 + 1 + 3 +
            (2 if i < len(questions) - 1 else (20 + (len(questions) - 1) * 2 + 1 + 2 + 1))
            for i, (_, q_title, _) in enumerate(questions)
        ]) + 13) + [0]
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in title] +
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in description]
        )  
        if len(questions) > 1:
            expected_send_calls += [call("+{TAB}", 0)] * 8 + [call("{ENTER}", 0)]
        if len(questions) > 2:
            expected_send_calls += ([call("+{TAB}", 0)] * 10 + [call("{ENTER}", 0)]) * (len(questions) - 2)
        for i, (question_type, question_title, is_required) in enumerate(questions):
            expected_send_calls += (
                [call(char, 1) for char in question_title] +
                [call("{TAB}", 0)] * 3 +
                [call("{ENTER}", 0)] +
                [call("{UP}{UP}", 0) if question_type == "Short answer" else call("{UP}", 0)] +
                [call("{ENTER}", 0)] +
                [call("{TAB}", 0)] * 3 +
                ([call("{SPACE}", 0)] if is_required else []) +
                (([call("{TAB}", 0)] * 2) if i < len(questions) - 1 else [])
            )
        expected_send_calls += (
            [call("+{TAB}", 0)] * (20 + (len(questions) - 1) * 2) +
            [call("{ENTER}", 0)]
        )
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + len(title) + len(description) + len(questions) - 1 + sum(len(question[1]) for question in questions)))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + len(title) +
            3 + 1 + len(description) +
            1 + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + 1 +
            sum([
                len(q_title) +
                3 + 1 + 1 + 1 + 3 +
                (2 if i < len(questions) - 1 else (20 + (len(questions) - 1) * 2 + 1))
                for i, (_, q_title, _) in enumerate(questions)
            ])
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "could not send enter key to press publish button")
    
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_send_key_to_move_to_publish_confirmation_button(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (len(name) + len(title) + len(description) + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + len(questions[0][1]) + 
        sum([
            len(q_title) +
            3 + 1 + 1 + 1 + 3 +
            (2 if i < len(questions) - 1 else (20 + (len(questions) - 1) * 2 + 1 + 2 + 1))
            for i, (_, q_title, _) in enumerate(questions)
        ]) + 14) + [0]
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in title] +
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in description]
        )  
        if len(questions) > 1:
            expected_send_calls += [call("+{TAB}", 0)] * 8 + [call("{ENTER}", 0)]
        if len(questions) > 2:
            expected_send_calls += ([call("+{TAB}", 0)] * 10 + [call("{ENTER}", 0)]) * (len(questions) - 2)
        for i, (question_type, question_title, is_required) in enumerate(questions):
            expected_send_calls += (
                [call(char, 1) for char in question_title] +
                [call("{TAB}", 0)] * 3 +
                [call("{ENTER}", 0)] +
                [call("{UP}{UP}", 0) if question_type == "Short answer" else call("{UP}", 0)] +
                [call("{ENTER}", 0)] +
                [call("{TAB}", 0)] * 3 +
                ([call("{SPACE}", 0)] if is_required else []) +
                (([call("{TAB}", 0)] * 2) if i < len(questions) - 1 else [])
            )
        expected_send_calls += (
            [call("+{TAB}", 0)] * (20 + (len(questions) - 1) * 2) +
            [call("{ENTER}", 0)] +
            [call("{TAB}{TAB}", 0)]
        )
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + len(title) + len(description) + len(questions) - 1 + sum(len(question[1]) for question in questions)))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + len(title) +
            3 + 1 + len(description) +
            1 + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + 1 +
            sum([
                len(q_title) +
                3 + 1 + 1 + 1 + 3 +
                (2 if i < len(questions) - 1 else (20 + (len(questions) - 1) * 2 + 1 + 1))
                for i, (_, q_title, _) in enumerate(questions)
            ])
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "could not send tab key to move to publish confirmation button")
        
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_send_key_to_press_publish_confirmation_button(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (len(name) + len(title) + len(description) + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + len(questions[0][1]) + 
        sum([
            len(q_title) +
            3 + 1 + 1 + 1 + 3 +
            (2 if i < len(questions) - 1 else (20 + (len(questions) - 1) * 2 + 1 + 2 + 1))
            for i, (_, q_title, _) in enumerate(questions)
        ]) + 15) + [0]
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in title] +
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in description]
        )  
        if len(questions) > 1:
            expected_send_calls += [call("+{TAB}", 0)] * 8 + [call("{ENTER}", 0)]
        if len(questions) > 2:
            expected_send_calls += ([call("+{TAB}", 0)] * 10 + [call("{ENTER}", 0)]) * (len(questions) - 2)
        for i, (question_type, question_title, is_required) in enumerate(questions):
            expected_send_calls += (
                [call(char, 1) for char in question_title] +
                [call("{TAB}", 0)] * 3 +
                [call("{ENTER}", 0)] +
                [call("{UP}{UP}", 0) if question_type == "Short answer" else call("{UP}", 0)] +
                [call("{ENTER}", 0)] +
                [call("{TAB}", 0)] * 3 +
                ([call("{SPACE}", 0)] if is_required else []) +
                (([call("{TAB}", 0)] * 2) if i < len(questions) - 1 else [])
            )
        expected_send_calls += (
            [call("+{TAB}", 0)] * (20 + (len(questions) - 1) * 2) +
            [call("{ENTER}", 0)] +
            [call("{TAB}{TAB}", 0)] +
            [call("{ENTER}", 0)]
        )
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + len(title) + len(description) + len(questions) - 1 + sum(len(question[1]) for question in questions)))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + len(title) +
            3 + 1 + len(description) +
            1 + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + 1 +
            sum([
                len(q_title) +
                3 + 1 + 1 + 1 + 3 +
                (2 if i < len(questions) - 1 else (20 + (len(questions) - 1) * 2 + 1 + 2))
                for i, (_, q_title, _) in enumerate(questions)
            ])
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "could not send enter key to press publish confirmation button")
        
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_send_tab_key_to_move_to_link_button(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (len(name) + len(title) + len(description) + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + len(questions[0][1]) + 
        sum([
            len(q_title) +
            3 + 1 + 1 + 1 + 3 +
            (2 if i < len(questions) - 1 else (20 + (len(questions) - 1) * 2 + 1 + 2 + 1))
            for i, (_, q_title, _) in enumerate(questions)
        ]) + 17) + [0]
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in title] +
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in description]
        )  
        if len(questions) > 1:
            expected_send_calls += [call("+{TAB}", 0)] * 8 + [call("{ENTER}", 0)]
        if len(questions) > 2:
            expected_send_calls += ([call("+{TAB}", 0)] * 10 + [call("{ENTER}", 0)]) * (len(questions) - 2)
        for i, (question_type, question_title, is_required) in enumerate(questions):
            expected_send_calls += (
                [call(char, 1) for char in question_title] +
                [call("{TAB}", 0)] * 3 +
                [call("{ENTER}", 0)] +
                [call("{UP}{UP}", 0) if question_type == "Short answer" else call("{UP}", 0)] +
                [call("{ENTER}", 0)] +
                [call("{TAB}", 0)] * 3 +
                ([call("{SPACE}", 0)] if is_required else []) +
                (([call("{TAB}", 0)] * 2) if i < len(questions) - 1 else [])
            )
        expected_send_calls += (
            [call("+{TAB}", 0)] * (20 + (len(questions) - 1) * 2) +
            [call("{ENTER}", 0)] +
            [call("{TAB}{TAB}", 0)] +
            [call("{ENTER}", 0)] +
            [call("+{TAB}", 0)] * 2
        )
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + len(title) + len(description) + len(questions) - 1 + sum(len(question[1]) for question in questions)))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + len(title) +
            3 + 1 + len(description) +
            1 + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + 1 +
            sum([
                len(q_title) +
                3 + 1 + 1 + 1 + 3 +
                (2 if i < len(questions) - 1 else (20 + (len(questions) - 1) * 2 + 1 + 2 + 1))
                for i, (_, q_title, _) in enumerate(questions)
            ]) +
            2
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "could not send tab key to move to link button")
        
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_send_space_key_to_move_to_link_button(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (len(name) + len(title) + len(description) + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + len(questions[0][1]) + 
        sum([
            len(q_title) +
            3 + 1 + 1 + 1 + 3 +
            (2 if i < len(questions) - 1 else (20 + (len(questions) - 1) * 2 + 1 + 2 + 1))
            for i, (_, q_title, _) in enumerate(questions)
        ]) + 18) + [0]
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in title] +
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in description]
        )  
        if len(questions) > 1:
            expected_send_calls += [call("+{TAB}", 0)] * 8 + [call("{ENTER}", 0)]
        if len(questions) > 2:
            expected_send_calls += ([call("+{TAB}", 0)] * 10 + [call("{ENTER}", 0)]) * (len(questions) - 2)
        for i, (question_type, question_title, is_required) in enumerate(questions):
            expected_send_calls += (
                [call(char, 1) for char in question_title] +
                [call("{TAB}", 0)] * 3 +
                [call("{ENTER}", 0)] +
                [call("{UP}{UP}", 0) if question_type == "Short answer" else call("{UP}", 0)] +
                [call("{ENTER}", 0)] +
                [call("{TAB}", 0)] * 3 +
                ([call("{SPACE}", 0)] if is_required else []) +
                (([call("{TAB}", 0)] * 2) if i < len(questions) - 1 else [])
            )
        expected_send_calls += (
            [call("+{TAB}", 0)] * (20 + (len(questions) - 1) * 2) +
            [call("{ENTER}", 0)] +
            [call("{TAB}{TAB}", 0)] +
            [call("{ENTER}", 0)] +
            [call("+{TAB}", 0)] * 2 +
            [call("{SPACE}", 0)]
        )
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + len(title) + len(description) + len(questions) - 1 + sum(len(question[1]) for question in questions)))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + len(title) +
            3 + 1 + len(description) +
            1 + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + 1 +
            sum([
                len(q_title) +
                3 + 1 + 1 + 1 + 3 +
                (2 if i < len(questions) - 1 else (20 + (len(questions) - 1) * 2 + 1 + 2 + 1))
                for i, (_, q_title, _) in enumerate(questions)
            ]) +
            2 + 1
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "could not send space key to move to link button")
         
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_send_key_to_move_to_shorten_link_checkbox(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (len(name) + len(title) + len(description) + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + len(questions[0][1]) + 
        sum([
            len(q_title) +
            3 + 1 + 1 + 1 + 3 +
            (2 if i < len(questions) - 1 else (20 + (len(questions) - 1) * 2 + 1 + 2 + 1))
            for i, (_, q_title, _) in enumerate(questions)
        ]) + 19) + [0]
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in title] +
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in description]
        )  
        if len(questions) > 1:
            expected_send_calls += [call("+{TAB}", 0)] * 8 + [call("{ENTER}", 0)]
        if len(questions) > 2:
            expected_send_calls += ([call("+{TAB}", 0)] * 10 + [call("{ENTER}", 0)]) * (len(questions) - 2)
        for i, (question_type, question_title, is_required) in enumerate(questions):
            expected_send_calls += (
                [call(char, 1) for char in question_title] +
                [call("{TAB}", 0)] * 3 +
                [call("{ENTER}", 0)] +
                [call("{UP}{UP}", 0) if question_type == "Short answer" else call("{UP}", 0)] +
                [call("{ENTER}", 0)] +
                [call("{TAB}", 0)] * 3 +
                ([call("{SPACE}", 0)] if is_required else []) +
                (([call("{TAB}", 0)] * 2) if i < len(questions) - 1 else [])
            )
        expected_send_calls += (
            [call("+{TAB}", 0)] * (20 + (len(questions) - 1) * 2) +
            [call("{ENTER}", 0)] +
            [call("{TAB}{TAB}", 0)] +
            [call("{ENTER}", 0)] +
            [call("+{TAB}", 0)] * 2 +
            [call("{SPACE}", 0)] +
            [call("{TAB}", 0)]
        )
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + len(title) + len(description) + len(questions) - 1 + sum(len(question[1]) for question in questions)))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + len(title) +
            3 + 1 + len(description) +
            1 + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + 1 +
            sum([
                len(q_title) +
                3 + 1 + 1 + 1 + 3 +
                (2 if i < len(questions) - 1 else (20 + (len(questions) - 1) * 2 + 1 + 2 + 1))
                for i, (_, q_title, _) in enumerate(questions)
            ]) +
            2 + 1 + 1
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "could not send tab key to move to shorten link checkbox")
        
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_send_key_to_check_shorten_link_checkbox(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (len(name) + len(title) + len(description) + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + len(questions[0][1]) + 
        sum([
            len(q_title) +
            3 + 1 + 1 + 1 + 3 +
            (2 if i < len(questions) - 1 else (20 + (len(questions) - 1) * 2 + 1 + 2 + 1))
            for i, (_, q_title, _) in enumerate(questions)
        ]) + 20) + [0]
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in title] +
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in description]
        )  
        if len(questions) > 1:
            expected_send_calls += [call("+{TAB}", 0)] * 8 + [call("{ENTER}", 0)]
        if len(questions) > 2:
            expected_send_calls += ([call("+{TAB}", 0)] * 10 + [call("{ENTER}", 0)]) * (len(questions) - 2)
        for i, (question_type, question_title, is_required) in enumerate(questions):
            expected_send_calls += (
                [call(char, 1) for char in question_title] +
                [call("{TAB}", 0)] * 3 +
                [call("{ENTER}", 0)] +
                [call("{UP}{UP}", 0) if question_type == "Short answer" else call("{UP}", 0)] +
                [call("{ENTER}", 0)] +
                [call("{TAB}", 0)] * 3 +
                ([call("{SPACE}", 0)] if is_required else []) +
                (([call("{TAB}", 0)] * 2) if i < len(questions) - 1 else [])
            )
        expected_send_calls += (
            [call("+{TAB}", 0)] * (20 + (len(questions) - 1) * 2) +
            [call("{ENTER}", 0)] +
            [call("{TAB}{TAB}", 0)] +
            [call("{ENTER}", 0)] +
            [call("+{TAB}", 0)] * 2 +
            [call("{SPACE}", 0)] +
            [call("{TAB}", 0)] +
            [call("{SPACE}", 0)]
        )
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + len(title) + len(description) + len(questions) - 1 + sum(len(question[1]) for question in questions)))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + len(title) +
            3 + 1 + len(description) +
            1 + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + 1 +
            sum([
                len(q_title) +
                3 + 1 + 1 + 1 + 3 +
                (2 if i < len(questions) - 1 else (20 + (len(questions) - 1) * 2 + 1 + 2 + 1))
                for i, (_, q_title, _) in enumerate(questions)
            ]) +
            2 + 1 + 1 + 1
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "could not send space key to check shorten link checkbox")
        
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_send_key_to_move_to_copy_link_button(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (len(name) + len(title) + len(description) + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + len(questions[0][1]) + 
        sum([
            len(q_title) +
            3 + 1 + 1 + 1 + 3 +
            (2 if i < len(questions) - 1 else (20 + (len(questions) - 1) * 2 + 1 + 2 + 1))
            for i, (_, q_title, _) in enumerate(questions)
        ]) + 22) + [0]
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in title] +
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in description]
        )  
        if len(questions) > 1:
            expected_send_calls += [call("+{TAB}", 0)] * 8 + [call("{ENTER}", 0)]
        if len(questions) > 2:
            expected_send_calls += ([call("+{TAB}", 0)] * 10 + [call("{ENTER}", 0)]) * (len(questions) - 2)
        for i, (question_type, question_title, is_required) in enumerate(questions):
            expected_send_calls += (
                [call(char, 1) for char in question_title] +
                [call("{TAB}", 0)] * 3 +
                [call("{ENTER}", 0)] +
                [call("{UP}{UP}", 0) if question_type == "Short answer" else call("{UP}", 0)] +
                [call("{ENTER}", 0)] +
                [call("{TAB}", 0)] * 3 +
                ([call("{SPACE}", 0)] if is_required else []) +
                (([call("{TAB}", 0)] * 2) if i < len(questions) - 1 else [])
            )
        expected_send_calls += (
            [call("+{TAB}", 0)] * (20 + (len(questions) - 1) * 2) +
            [call("{ENTER}", 0)] +
            [call("{TAB}{TAB}", 0)] +
            [call("{ENTER}", 0)] +
            [call("+{TAB}", 0)] * 2 +
            [call("{SPACE}", 0)] +
            [call("{TAB}", 0)] +
            [call("{SPACE}", 0)] +
            [call("{TAB}", 0)] * 2
        )
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + len(title) + len(description) + len(questions) - 1 + sum(len(question[1]) for question in questions)))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + len(title) +
            3 + 1 + len(description) +
            1 + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + 1 +
            sum([
                len(q_title) +
                3 + 1 + 1 + 1 + 3 +
                (2 if i < len(questions) - 1 else (20 + (len(questions) - 1) * 2 + 1 + 2 + 1))
                for i, (_, q_title, _) in enumerate(questions)
            ]) +
            2 + 1 + 1 + 1 + 1 + 2
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "could not send tab key to move to copy link button")
    
    @patch("time.sleep", return_value=None)
    def test_create_form_failed_to_send_key_to_copy_link_to_clipboard(self, mock_sleep):
        name = "Example Form Name"
        title = "Example Form Title"
        description = "Example form description"
        questions = [
            ["Short answer", "question 1", True],
            ["Paragraph", "question 2", True],
            ["Short answer", "question 3", False]
        ]
        self.gforms.browser.check_existing_window.return_value = "Browser window didn't exist"
        self.gforms.browser.create_window.return_value = None
        self.gforms.browser.browse.return_value = None
        self.gforms.browser.toggle_fullscreen.return_value = None
        self.gforms.dll.AU3_MouseClick.return_value = 1
        self.gforms.dll.AU3_Send.side_effect = [1] * (len(name) + len(title) + len(description) + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + len(questions[0][1]) + 
        sum([
            len(q_title) +
            3 + 1 + 1 + 1 + 3 +
            (2 if i < len(questions) - 1 else (20 + (len(questions) - 1) * 2 + 1 + 2 + 1))
            for i, (_, q_title, _) in enumerate(questions)
        ]) + 23) + [0]
        self.gforms.dll.AU3_WinActive.return_value = 1

        res = self.gforms.create_form(name=name, title=title, description=description, questions=questions)
        
        self.gforms.browser.check_existing_window.assert_called_once()
        self.gforms.browser.create_window.assert_called_once()
        self.gforms.browser.create_tab.assert_not_called()
        self.gforms.browser.browse.assert_called_once_with("https://forms.google.com/create")
        self.gforms.browser.toggle_fullscreen.assert_called_once()
        self.gforms.dll.AU3_MouseClick.assert_called_once_with("left", 0, 0, 1, 10)
        expected_send_calls = (
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in name] +
            [call("{TAB}", 0)] * 18 +
            [call("^a{BACKSPACE}", 0)] +
            [call(char, 1) for char in title] +
            [call("{TAB}", 0)] * 2 +
            [call(char, 1) for char in description]
        )  
        if len(questions) > 1:
            expected_send_calls += [call("+{TAB}", 0)] * 8 + [call("{ENTER}", 0)]
        if len(questions) > 2:
            expected_send_calls += ([call("+{TAB}", 0)] * 10 + [call("{ENTER}", 0)]) * (len(questions) - 2)
        for i, (question_type, question_title, is_required) in enumerate(questions):
            expected_send_calls += (
                [call(char, 1) for char in question_title] +
                [call("{TAB}", 0)] * 3 +
                [call("{ENTER}", 0)] +
                [call("{UP}{UP}", 0) if question_type == "Short answer" else call("{UP}", 0)] +
                [call("{ENTER}", 0)] +
                [call("{TAB}", 0)] * 3 +
                ([call("{SPACE}", 0)] if is_required else []) +
                (([call("{TAB}", 0)] * 2) if i < len(questions) - 1 else [])
            )
        expected_send_calls += (
            [call("+{TAB}", 0)] * (20 + (len(questions) - 1) * 2) +
            [call("{ENTER}", 0)] +
            [call("{TAB}{TAB}", 0)] +
            [call("{ENTER}", 0)] +
            [call("+{TAB}", 0)] * 2 +
            [call("{SPACE}", 0)] +
            [call("{TAB}", 0)] +
            [call("{SPACE}", 0)] +
            [call("{TAB}", 0)] * 2 +
            [call("{SPACE}", 0)]
        )
        self.gforms.dll.AU3_Send.assert_has_calls(expected_send_calls)
        self.gforms.dll.AU3_WinActive.assert_has_calls([call(self.gforms.browser.window_info, "")] * (len(name) + len(title) + len(description) + len(questions) - 1 + sum(len(question[1]) for question in questions)))
        
        expected_sleep_calls = (
            5 + 1 + len(name) +
            19 + 1 + len(title) +
            3 + 1 + len(description) +
            1 + (8 if len(questions) > 1 else 0) + (10 * (len(questions) - 2) if len(questions) > 2 else 0) + 1 +
            sum([
                len(q_title) +
                3 + 1 + 1 + 1 + 3 +
                (2 if i < len(questions) - 1 else (20 + (len(questions) - 1) * 2 + 1 + 2 + 1))
                for i, (_, q_title, _) in enumerate(questions)
            ]) +
            2 + 1 + 1 + 1 + 1 + 2 + 1
        )
        self.assertEqual(mock_sleep.call_count, expected_sleep_calls)
        self.assertEqual(res, "could not send space key to copy link to clipboard")