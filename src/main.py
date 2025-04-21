import yaml
import time
import random
import ctypes
import sys
from activities.apps.browser_apps.google_forms import GoogleForms
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
    "Google Forms": GoogleForms,
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

def get_native_screen_resolution():
    logger.info("Checking native screen resolution")
    user32 = ctypes.windll.user32
    logger.info("Setting process to DPI aware")
    user32.SetProcessDPIAware()
    screen_width = user32.GetSystemMetrics(0)
    screen_height = user32.GetSystemMetrics(1)
    return screen_width, screen_height
    

def load_config(file_path="configs/scenario.yaml"):
    logger.info("Loading config file scenario.yaml")
    with open(file_path, "r") as file:
        return yaml.safe_load(file)

if __name__ == "__main__":
    screen_width, screen_height = get_native_screen_resolution()
    if screen_width != 1920 or screen_height != 1080:
        logger.error("Native screen resolution is not 1920x1080")
        sys.exit(1)

    config = load_config()

    execution_mode = config.get("execution_mode", "sequential")
    repeat = config.get("repeat", False)
    scenarios = config.get("scenario", [])

    while True:
        if execution_mode == "random":
            random.shuffle(scenarios)

        for s in scenarios:
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
                    except Exception as e:
                        logger.error(f"Exception while running {method} on {name}: {e}")
                else:
                    logger.error(f"Method '{method}' not found on '{name}'")

                time.sleep(delay)

        if not repeat:
            break