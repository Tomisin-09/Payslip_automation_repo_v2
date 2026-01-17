from pathlib import Path
import logging
from datetime import datetime


def setup_logging(
    base_output_dir: Path,
    period_id: str,
    logging_cfg: dict,
) -> logging.Logger:
    """
    Configure logging with:
    - Console output
    - Timestamped log files
    - Daily log directories
    """

    # --------------------------------------------------
    # Resolve timestamps
    # --------------------------------------------------
    now = datetime.now()
    date_dir = now.strftime("%Y-%m-%d")
    timestamp = now.strftime("%Y%m%d_%H%M%S")

    # --------------------------------------------------
    # Resolve log directory
    # --------------------------------------------------
    logs_base_dir = (
        base_output_dir
        / period_id
        / logging_cfg.get("logs_dir", "logs")
        / date_dir
    )
    logs_base_dir.mkdir(parents=True, exist_ok=True)

    log_file_path = logs_base_dir / f"run_{timestamp}.log"

    # --------------------------------------------------
    # Root logger
    # --------------------------------------------------
    logger = logging.getLogger()
    logger.setLevel(logging_cfg.get("level", "INFO"))

    # Clear existing handlers (important for reruns)
    logger.handlers.clear()

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)-5s | %(name)s | %(message)s"
    )

    # --------------------------------------------------
    # File handler
    # --------------------------------------------------
    if logging_cfg.get("log_to_file", True):
        file_handler = logging.FileHandler(log_file_path)
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)

    # --------------------------------------------------
    # Console handler
    # --------------------------------------------------
    if logging_cfg.get("log_to_console", True):
        console_handler = logging.StreamHandler()
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)

    logger.info("Logging initialised")
    logger.info("Log file: %s", log_file_path)

    return logger
