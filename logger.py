"""
Logging module for Certificate Generator
Provides centralized logging functionality
"""

import logging
import sys
from datetime import datetime
from pathlib import Path
from typing import Optional
from config import config_manager

class ColoredFormatter(logging.Formatter):
    """Custom formatter with colors for console output"""
    
    COLORS = {
        'DEBUG': '\033[36m',    # Cyan
        'INFO': '\033[32m',     # Green
        'WARNING': '\033[33m',  # Yellow
        'ERROR': '\033[31m',    # Red
        'CRITICAL': '\033[35m', # Magenta
        'RESET': '\033[0m'      # Reset
    }
    
    def format(self, record):
        log_color = self.COLORS.get(record.levelname, self.COLORS['RESET'])
        reset_color = self.COLORS['RESET']
        
        # Add color to the level name
        record.levelname = f"{log_color}{record.levelname}{reset_color}"
        return super().format(record)

def setup_logger(name: str = "certificate_generator", 
                log_file: Optional[str] = None,
                level: Optional[str] = None) -> logging.Logger:
    """
    Set up logger with both console and file output
    
    Args:
        name: Logger name
        log_file: Path to log file (optional)
        level: Log level (optional, uses config if not provided)
    
    Returns:
        Configured logger instance
    """
    logger = logging.getLogger(name)
    
    # Clear existing handlers
    logger.handlers.clear()
    
    # Set level
    if level is None:
        level = config_manager.get_config().log_level
    logger.setLevel(getattr(logging, level.upper(), logging.INFO))
    
    # Create formatters
    console_formatter = ColoredFormatter(
        '%(asctime)s [%(levelname)s] %(name)s: %(message)s',
        datefmt='%H:%M:%S'
    )
    
    file_formatter = logging.Formatter(
        '%(asctime)s [%(levelname)s] %(name)s: %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(console_formatter)
    logger.addHandler(console_handler)
    
    # File handler (if specified)
    if log_file:
        log_path = Path(log_file)
        log_path.parent.mkdir(parents=True, exist_ok=True)
        
        file_handler = logging.FileHandler(log_path, encoding='utf-8')
        file_handler.setFormatter(file_formatter)
        logger.addHandler(file_handler)
    
    return logger

def get_logger(name: str = "certificate_generator") -> logging.Logger:
    """Get logger instance"""
    return logging.getLogger(name)

# Create default logger
default_logger = setup_logger(log_file="logs/certificate_generator.log")