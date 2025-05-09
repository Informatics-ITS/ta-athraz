import unittest
from unittest.mock import patch, MagicMock
from activities.others.selenium_script import SeleniumScript

class TestSeleniumScript(unittest.TestCase):
    def setUp(self):
        self.selenium_script = SeleniumScript()
    
    @patch("os.path.exists", return_value=True)
    @patch("subprocess.run")
    def test_run_pytest_success(self, mock_run, mock_exists):
        path = "path/to/testcase.py"
        mock_process = MagicMock()
        mock_process.returncode = 0
        mock_process.stdout = "Test passed output\n"
        mock_process.stderr = ""
        mock_run.return_value = mock_process

        res = self.selenium_script.run_pytest(path=path)
        
        mock_exists.assert_called_once_with(path)
        mock_run.assert_called_once_with(["pytest", path, "-v"], capture_output=True, text=True)
        self.assertIsNone(res)
        
    @patch("os.path.exists", return_value=False)
    def test_run_pytest_failed_path_doesnt_exist(self, mock_exists):
        path = "path/to/testcase.py"
        res = self.selenium_script.run_pytest(path=path)
        
        mock_exists.assert_called_once_with(path)
        self.assertEqual(res, f"Script file '{path}' does not exist")
        
    @patch("os.path.exists", return_value=True)
    @patch("subprocess.run")
    def test_run_pytest_failed_on_run(self, mock_run, mock_exists):
        path = "path/to/testcase.py"
        mock_process = MagicMock()
        mock_process.returncode = 1
        mock_process.stdout = "Some tests failed\n"
        mock_process.stderr = "Error info\n"
        mock_run.return_value = mock_process

        res = self.selenium_script.run_pytest(path=path)
        
        mock_exists.assert_called_once_with(path)
        mock_run.assert_called_once_with(["pytest", path, "-v"], capture_output=True, text=True)
        expected_output = mock_process.stdout + mock_process.stderr
        self.assertEqual(res, expected_output.strip())