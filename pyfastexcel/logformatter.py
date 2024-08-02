import logging
import warnings

NORMAL_FORMAT = '\033[%(log_color)s%(asctime)s - %(name)s - %(levelname)s - %(message)s\033[0m'
LOG_COLORS = {
    'DEBUG': '\033[90m',  # Grey
    'INFO': '\033[92m',  # Green
    'WARNING': '\033[93m',  # Yellow
    'ERROR': '\033[91m',  # Red
    'CRITICAL': '\033[91m',  # Red
}


def custom_warning_format(message, *args):
    start_color = '\033[93m'
    reset_color = '\033[0m'
    return f'{start_color}DeprecationWarning: {message}{reset_color}\n'


warnings.formatwarning = custom_warning_format
warnings.simplefilter('always', DeprecationWarning)


class ColoredFormatter(logging.Formatter):
    def __init__(self, fmt):
        super().__init__(fmt)

    def format(self, record):
        record.log_color = LOG_COLORS[record.levelname]
        return super().format(record)


formatter = ColoredFormatter(NORMAL_FORMAT)
