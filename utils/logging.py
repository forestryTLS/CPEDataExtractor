import logging
from enum import StrEnum

class bcolors(StrEnum):
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

level_to_color_map = {
    logging.INFO: bcolors.OKCYAN,
    logging.WARNING: bcolors.WARNING,
    logging.ERROR: bcolors.FAIL,
    logging.CRITICAL: bcolors.HEADER
}

class CustomStreamHandler(logging.Handler):
    CYAN = '\033[96m'
    YELLOW = '\033[93m'
    RED = '\033[91m'
    PURPLE = '\033[95m'
    GREEN = '\033[92m'
    ENDC = '\033[0m'

    LEVEL_TO_COLOR_MAP = {
        logging.DEBUG: GREEN,
        logging.INFO: CYAN,
        logging.WARNING: YELLOW,
        logging.ERROR: RED,
        logging.CRITICAL: PURPLE
    }

    def emit(self, record):
        """ Sends the formatted record message to stdout, styled with a terminal color that matches the record level.  """
        color_start = (
            self.LEVEL_TO_COLOR_MAP[record.levelno] 
            if record.levelno in self.LEVEL_TO_COLOR_MAP
            else ''
        )

        color_end = (
            self.ENDC
            if record.levelno in self.LEVEL_TO_COLOR_MAP
            else ''
        )

        print('{0}{1}{2}'.format(
            color_start,
            self.format(record),
            color_end
        ))
        
def print_colored(*args, color: bcolors):
    """ Print text to the terminal in the specified `color`. """
    if not isinstance(color, bcolors):
        raise TypeError("The argument \"color\" must be of type \"bcolors\".")
    
    print(color, end='')
    print(*args, end='')
    print(bcolors.ENDC)