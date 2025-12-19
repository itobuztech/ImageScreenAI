# Dependencies
import sys
import logging
from datetime import datetime
from config.settings import settings


class ColoredFormatter(logging.Formatter):
    """
    Colored log formatter for better readability
    """
    COLORS = {'DEBUG'    : '\033[36m',  # Cyan
              'INFO'     : '\033[32m',  # Green
              'WARNING'  : '\033[33m',  # Yellow
              'ERROR'    : '\033[31m',  # Red
              'CRITICAL' : '\033[35m',  # Magenta
              'RESET'    : '\033[0m',
             }
    

    def format(self, record):
        if sys.stdout.isatty():
            levelname = record.levelname

            if (levelname in self.COLORS):
                record.levelname = f"{self.COLORS[levelname]}{levelname}{self.COLORS['RESET']}"
        
        return super().format(record)


def setup_logger(name: str = None) -> logging.Logger:
    """
    Setup logger with console and file handlers
    
    Arguments:
    ----------
        name   { str }     : Logger name (defaults to root logger)
    
    Returns:
    --------
        { logging.Logger } : Configured logger instance
    """
    logger = logging.getLogger(name or settings.APP_NAME)
    
    # Avoid duplicate handlers
    if logger.handlers:
        return logger
    
    level             = getattr(logging, settings.LOG_LEVEL, logging.INFO)
    logger.setLevel(level)

    logger.propagate  = False
    
    # Console handler with colors
    console_handler   = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.DEBUG if settings.DEBUG else logging.INFO)
    
    console_formatter = ColoredFormatter('%(asctime)s | %(levelname)-8s | %(name)s | %(message)s',
                                         datefmt = '%Y-%m-%d %H:%M:%S'
                                        )
    console_handler.setFormatter(console_formatter)

    logger.addHandler(console_handler)
    
    # File handler
    log_file          = settings.LOGS_DIR / f"app_{datetime.now().strftime('%Y%m%d')}.log"
    file_handler      = logging.FileHandler(log_file)
    file_handler.setLevel(logging.DEBUG)

    file_formatter    = logging.Formatter('%(asctime)s | %(levelname)-8s | %(name)s | %(funcName)s:%(lineno)d | %(message)s',
                                          datefmt = '%Y-%m-%d %H:%M:%S'
                                         )

    file_handler.setFormatter(file_formatter)
    
    logger.addHandler(file_handler)
    
    return logger


def get_logger(name: str = None) -> logging.Logger:
    """
    Get or create logger instance
    """
    return setup_logger(name)