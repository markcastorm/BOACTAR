"""
BOACTAR Logging Configuration
=============================
Sets up logging for all modules in the BOACTAR runbook.
"""

import logging
import os
from datetime import datetime

import config


def setup_logging(timestamp=None):
    """
    Configure logging for the BOACTAR runbook.

    Creates both file and console handlers with unified formatting.
    Log files are stored in timestamped subdirectories.

    Args:
        timestamp: Optional timestamp string. If None, generates current timestamp.

    Returns:
        str: Path to the log file created.
    """
    # Generate timestamp if not provided
    if timestamp is None:
        timestamp = config.get_timestamp()

    # Create logs directory structure
    log_dir = os.path.join(config.LOGS_DIR, timestamp)
    os.makedirs(log_dir, exist_ok=True)

    # Also create latest symlink/directory
    latest_dir = os.path.join(config.LOGS_DIR, config.LATEST_FOLDER)
    os.makedirs(latest_dir, exist_ok=True)

    # Log file paths
    log_filename = f'boactar_{timestamp}.log'
    log_file = os.path.join(log_dir, log_filename)
    latest_log_file = os.path.join(latest_dir, 'boactar_latest.log')

    # Determine log level
    log_level = logging.DEBUG if config.DEBUG_MODE else getattr(logging, config.LOG_LEVEL, logging.INFO)

    # Configure root logger
    root_logger = logging.getLogger()
    root_logger.setLevel(log_level)

    # Clear any existing handlers
    root_logger.handlers = []

    # Create formatter
    formatter = logging.Formatter(config.LOG_FORMAT)

    # File handler for timestamped log
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(log_level)
    file_handler.setFormatter(formatter)
    root_logger.addHandler(file_handler)

    # File handler for latest log (overwrites)
    latest_handler = logging.FileHandler(latest_log_file, mode='w', encoding='utf-8')
    latest_handler.setLevel(log_level)
    latest_handler.setFormatter(formatter)
    root_logger.addHandler(latest_handler)

    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)
    console_handler.setFormatter(formatter)
    root_logger.addHandler(console_handler)

    # Log initialization
    logger = logging.getLogger(__name__)
    logger.info(f'Logging initialized - Level: {config.LOG_LEVEL}')
    logger.info(f'Log file: {log_file}')

    return log_file


def get_logger(name):
    """
    Get a logger instance for a module.

    Args:
        name: Module name (typically __name__).

    Returns:
        logging.Logger: Configured logger instance.
    """
    return logging.getLogger(name)


if __name__ == '__main__':
    # Test logging setup
    log_file = setup_logging()
    logger = get_logger(__name__)

    logger.debug('Debug message test')
    logger.info('Info message test')
    logger.warning('Warning message test')
    logger.error('Error message test')

    print(f'\nLog file created: {log_file}')
