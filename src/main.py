import yaml
import time
import random
from scenarios.chrome_scenario import ChromeScenario
from scenarios.word_scenario import WordScenario
from configs.logger import logger

def load_config(file_path = "configs/scenarios.yaml"):
    with open(file_path, "r") as file:
        return yaml.safe_load(file)

scenario_map = {
    "chrome_scenario": ChromeScenario,
    "word_scenario": WordScenario,
}

if __name__ == "__main__":
    config = load_config()

    execution_mode = config.get("execution_mode", "sequential")
    repeat = config.get("repeat", True)
    scenarios = config.get("scenarios", [])

    while True:
        if execution_mode == "random":
            random.shuffle(scenarios)
            
        logger.info(scenarios)

        for s in scenarios:
            name = s["name"]
            method = s.get("method", "run")
            delay = s.get("delay", 0)
            args = s.get("args", [])

            if name in scenario_map:
                scenario_instance = scenario_map[name]()
                run_method = getattr(scenario_instance, method, None)
                if run_method:
                    logger.info(f"Running method {method} on {name}")
                    err = run_method(*args)
                    if err:
                        logger.error(f"Error running method {method} on {name}: {err}")
                else:
                    logger.error(f"Method {method} not found on {name}")

            else:
                logger.error(f"Scenario '{name}' not found in scenario map")

            time.sleep(delay)

        if not repeat:
            break