import os
import subprocess
from configs.logger import logger

class SeleniumScript():
    def run_pytest(self, path):
        if not os.path.exists(path):
            return f"Script file '{path}' does not exist"

        result = subprocess.run(
            ["pytest", path, "-v"],
            capture_output=True,
            text=True
        )

        output = result.stdout + result.stderr
        if output.strip():
            logger.info(output)

        if result.returncode != 0:
            return output.strip()
            
        return None