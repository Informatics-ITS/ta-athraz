import yaml
import time
import random
from activities.apps.browser_apps.google_forms import GoogleForms
from activities.apps.browsers.google_chrome import GoogleChrome
from activities.apps.browsers.mozilla_firefox import MozillaFirefox
from activities.apps.native_apps.microsoft_word import MicrosoftWord
from configs.logger import logger

def load_config(file_path="configs/scenario.yaml"):
    with open(file_path, "r") as file:
        return yaml.safe_load(file)

activity_map = {
    "Google Forms": GoogleForms,
    "Google Chrome": GoogleChrome,
    "Mozilla Firefox": MozillaFirefox,
    "Microsoft Word": MicrosoftWord,
}

if __name__ == "__main__":
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

            for m in methods:
                method = m.get("method")
                delay = m.get("delay", 0)
                args = m.get("args", {})

                if name in activity_map:
                    if browser:
                        browser_instance = activity_map[browser]()
                        activity_instance = activity_map[name](browser_instance)
                    else:
                        activity_instance = activity_map[name]()

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
                else:
                    logger.error(f"Activity '{name}' not found in activity map")

                time.sleep(delay)

        if not repeat:
            break