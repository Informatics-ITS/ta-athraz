import logging

log_format = "%(asctime)s - %(levelname)s - %(message)s"
logging.basicConfig(
    level = logging.DEBUG,
    format = log_format,
    datefmt = "%Y-%m-%d %H:%M:%S",
    handlers = [
        logging.FileHandler(f"app.log"),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger(__name__)