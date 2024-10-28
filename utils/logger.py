import logging
from pathlib import Path
from typing import Optional

def setup_logger(name: str, log_level: Optional[int] = None) -> logging.Logger:
    """
    Configure and return a logger instance.
    
    Args:
        name: Name for the logger
        log_level: Optional logging level (defaults to DEBUG)
        
    Returns:
        Configured logger instance
    """
    # Ensure logs directory exists
    logs_dir = Path("logs")
    logs_dir.mkdir(exist_ok=True)

    # Create logger
    logger = logging.getLogger(name)
    logger.setLevel(log_level or logging.DEBUG)

    # Remove any existing handlers
    logger.handlers = []

    # Create handlers
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)

    log_file_path = logs_dir / f'{name}.log'
    file_handler = logging.FileHandler(log_file_path, mode='a')
    file_handler.setLevel(logging.DEBUG)

    # Create formatter
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s'
    )

    # Set formatter for handlers
    console_handler.setFormatter(formatter)
    file_handler.setFormatter(formatter)

    # Add handlers to logger
    logger.addHandler(console_handler)
    logger.addHandler(file_handler)

    return logger
