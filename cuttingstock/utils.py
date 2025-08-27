import logging

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/cuttingstock.log'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)


def log_message(level: str, message: str, details: dict = None):
    """
    Log a message with the specified level.

    Args:
        level (str): The log level ('debug', 'info', 'warning', 'error', 'critical')
        message (str): The main message to log
        details (dict, optional): Additional details to include in the log
    """
    log_func = getattr(logger, level.lower(), logger.info)

    if details:
        detail_str = ", ".join([f"{k}: {v}" for k, v in details.items()])
        log_func(f"{message} | Details: {detail_str}")
    else:
        log_func(message)
