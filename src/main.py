import yaml
import time
import random
import sys
from ctypes import windll
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

def get_display_resolution_and_scale():
    logger.info("Checking display resolution and scale")
    user32 = windll.user32
    gdi32 = windll.gdi32
    logger.info("Setting process to DPI aware")
    user32.SetProcessDPIAware()
    screen_width = user32.GetSystemMetrics(0)
    screen_height = user32.GetSystemMetrics(1)
    hdc = user32.GetDC(0)
    dpi = gdi32.GetDeviceCaps(hdc, 88)
    user32.ReleaseDC(0, hdc)
    scale = int((dpi / 96) * 100)
    return screen_width, screen_height, scale

def load_config(file_path="configs/scenarios.yaml"):
    logger.info("Loading config file scenarios.yaml")
    with open(file_path, "r") as file:
        return yaml.safe_load(file)

if __name__ == "__main__":
    screen_width, screen_height, scale = get_display_resolution_and_scale()
    if screen_width != 1920 or screen_height != 1080:
        logger.error("Display resolution is not 1920x1080")
        sys.exit(1)
    if scale != 125:
        logger.error("Scale is not 125%")
        sys.exit(1)

    config = load_config()

    execution_mode = config.get("execution_mode", "sequential")
    repeat = config.get("repeat", False)
    scenarios = config.get("scenarios", [])

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
                        except Exception as e:
                            logger.error(f"Exception while running {method} on {name}: {e}")
                    else:
                        logger.error(f"Method '{method}' not found on '{name}'")

                    time.sleep(delay)

        if not repeat:
            break