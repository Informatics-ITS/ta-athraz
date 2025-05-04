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

def set_dpi_aware():
    logger.info("Setting process to DPI aware")
    user32 = windll.user32
    user32.SetProcessDPIAware()

def load_config(file_path="configs/scenarios.yaml"):
    logger.info("Loading config file scenarios.yaml")
    with open(file_path, "r") as file:
        return yaml.safe_load(file)

if __name__ == "__main__":
    set_dpi_aware()
    config = load_config()

    execution_mode = config.get("execution_mode", "sequential")
    repeat = config.get("repeat", False)
    exit_on_error = config.get("exit_on_error", True)
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