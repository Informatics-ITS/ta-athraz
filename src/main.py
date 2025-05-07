import yaml
import time
import random
import sys
import ctypes
from activities.apps.browser_apps.gmail import Gmail
from activities.apps.browser_apps.google_forms import GoogleForms
from activities.apps.browser_apps.youtube import YouTube
from activities.apps.browsers.google_chrome import GoogleChrome
from activities.apps.browsers.mozilla_firefox import MozillaFirefox
from activities.apps.native_apps.microsoft_excel import MicrosoftExcel
from activities.apps.native_apps.microsoft_paint import MicrosoftPaint
from activities.apps.native_apps.microsoft_word import MicrosoftWord
from activities.apps.native_apps.notepad import Notepad
from activities.files.command_prompt import CommandPrompt
from activities.files.file_explorer import FileExplorer
from activities.others.autoit_function import AutoItFunction
from activities.others.selenium_script import SeleniumScript
from configs.logger import logger

activity_map = {
    "Gmail": Gmail,
    "Google Forms": GoogleForms,
    "YouTube": YouTube,
    "Google Chrome": GoogleChrome,
    "Mozilla Firefox": MozillaFirefox,
    "Microsoft Excel": MicrosoftExcel,
    "Microsoft Paint": MicrosoftPaint,
    "Microsoft Word": MicrosoftWord,
    "Notepad": Notepad,
    "Command Prompt": CommandPrompt,
    "File Explorer": FileExplorer,
    "AutoIt Function": AutoItFunction,
    "Selenium Script": SeleniumScript,
}

activity_instances = {}

def set_dpi_aware():
    logger.info("Setting process to DPI aware")
    user32 = ctypes.windll.user32
    result = user32.SetProcessDPIAware()
    if result == 0:
        logger.error("Failed to set process to DPI aware")
    else:
        logger.info("Process set to DPI aware successfully")
    
def set_taskbar_autohide():
    logger.info("Setting taskbar to auto-hide")
    class APPBARDATA(ctypes.Structure):
        _fields_ = [
            ("cbSize", ctypes.c_uint),
            ("hWnd", ctypes.c_void_p),
            ("uCallbackMessage", ctypes.c_uint),
            ("uEdge", ctypes.c_uint),
            ("rc", ctypes.c_int * 4),
            ("lParam", ctypes.c_int)
        ]

    appbar_data = APPBARDATA()
    appbar_data.cbSize = ctypes.sizeof(APPBARDATA)
    appbar_data.lParam = 0x1

    result = ctypes.windll.shell32.SHAppBarMessage(0x0000000a, ctypes.byref(appbar_data))
    if result == 0:
        logger.error("Failed to set taskbar to auto-hide using SHAppBarMessage")
    else:
        logger.info("Taskbar auto-hide successfully enabled")

def load_config(file_path="configs/scenarios.yaml"):
    logger.info(f"Loading config file: {file_path}")
    
    try:
        with open(file_path, "r") as file:
            config = yaml.safe_load(file)
            logger.info("Scenarios file loaded successfully")
            return config
    except FileNotFoundError:
        logger.error(f"Scenarios file not found: {file_path}")
    except yaml.YAMLError as e:
        logger.error(f"Error parsing YAML file {file_path}: {e}")
    except Exception as e:
        logger.error(f"Unexpected error while loading scenarios file {file_path}: {e}")

if __name__ == "__main__":
    set_dpi_aware()
    set_taskbar_autohide()
    config = load_config()

    execution_mode = config.get("execution_mode", "sequential")
    repeat = config.get("repeat", False)
    exit_on_error = config.get("exit_on_error", True)
    scenarios = config.get("scenarios", [])

    time.sleep(1)
    while True:
        if execution_mode == "random":
            random.shuffle(scenarios)

        for scenario in scenarios:
            scenario = scenario.get("scenario", [])

            for s in scenario:
                name = s["name"]
                browser = s.get("browser", None)
                methods = s.get("methods", [])

                if name not in activity_instances:
                    if browser:
                        if browser not in activity_instances:
                            activity_instances[browser] = activity_map[browser]()
                        activity_instances[name] = activity_map[name](activity_instances[browser])
                    else:
                        activity_instances[name] = activity_map[name]()

                activity_instance = activity_instances[name]

                for m in methods:
                    method = m.get("method")
                    delay = m.get("delay", 0)
                    args = m.get("args", {})

                    run_method = getattr(activity_instance, method, None)
                    if run_method:
                        logger.info(f"Running method '{method}' on activity '{name}' with args: {args}")
                        try:
                            if isinstance(args, dict):
                                err = run_method(**args)
                            elif isinstance(args, list):
                                err = run_method(*args)
                            else:
                                err = run_method()

                            if err:
                                logger.error(f"Error running method {method} on {name}: {err}")
                                if exit_on_error:
                                    sys.exit(1)
                        except Exception as e:
                            logger.error(f"Exception while running {method} on {name}: {e}")
                            if exit_on_error:
                                    sys.exit(1)
                    else:
                        logger.error(f"Method '{method}' not found on '{name}'")
                        if exit_on_error:
                            sys.exit(1)

                    time.sleep(delay)

        if not repeat:
            break