import logging
import sys
from logging.handlers import TimedRotatingFileHandler
FORMATTER = logging.Formatter("%(asctime)s | %(name)s | %(levelname)s | %(funcName)s:%(lineno)d | %(message)s")
LOG_FILE = r"C:/Users/a248433/Documents/SB Mozambique/EDO/Robotics/Bots/Automatização do Processo Abertura de contas Corporate/logs/{}.log".format('logs_')

def get_console_handler():
   console_handler = logging.StreamHandler(sys.stdout)
   console_handler.setFormatter(FORMATTER)
   return console_handler

def get_file_handler():
   file_handler = TimedRotatingFileHandler(LOG_FILE, when='midnight')
   file_handler.setFormatter(FORMATTER)
   return file_handler

def get_logger(logger_name):
   logger = logging.getLogger(logger_name)
   logger.setLevel(logging.INFO) # better to have too much log than not enough
   if not logger.hasHandlers():
       logger.addHandler(get_console_handler())
       logger.addHandler(get_file_handler())
   # with this pattern, it's rarely necessary to propagate the error up to parent
   logger.propagate = False
   return logger

# ● DEBUG: You should use this level for debugging purposes in development.
# ● INFO: You should use this level when something interesting—but expected—happens (e.g., a user starts a new project in a project management application).
# ● WARNING: You should use this level when something unexpected or unusual happens. It’s not an error, but you should pay attention to it.
# ● ERROR: This level is for things that go wrong but are usually recoverable (e.g., internal exceptions you can handle or APIs returning error results).
# ● CRITICAL: You should use this level in a doomsday scenario. The application is unusable. At this level, someone should be woken up at 2 a.m.

