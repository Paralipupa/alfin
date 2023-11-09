import logging
import os
import logging.config
from module.logger import CustomFilter, CustomFormatter

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if not os.path.exists(os.path.join(BASE_DIR, "logs")):
    os.makedirs(os.path.join(BASE_DIR, "logs"))
PATH_DUMP = os.path.join(BASE_DIR, "dump")
if not os.path.exists(PATH_DUMP):
    os.makedirs(PATH_DUMP)
ERROR_LOG_FILENAME = os.path.join(BASE_DIR, "logs", "error.json")
DEBUG_LOG_FILENAME = os.path.join(BASE_DIR, "logs", "debug.log")
INFO_LOG_FILENAME = os.path.join(BASE_DIR, "logs", "info.log")


ENCONING = "utf-8"
PATT_NAME = "(?:[А-ЯЁ][А-Яа-яё-]+\s+){2}((?:[А-ЯЁ][А-Яа-яё-]*\s*){1,2}(?:на|ва|ич|глы|глу|гли|улу|ызы)|-)"
PATT_NAME_ALT = "Центральный офис"
# PATT_NAME = "(?:вна|вич|глы|глу|гли|улу|ызы)\s*$"
PATT_FAMALY = "^\w+"
PATT_CURRENCY = "^-?\d{1,8}(?:[\.,]\d+)?$"
PATT_RATE = "^\d{1,3}(?:[\.,]\d+)?$"
PATT_PDN = "^\d{1,5}(?:[\.,]\d+)?$"
PATT_TARIF = "(?:постоянный|старт|31|24)$"
PATT_PERIOD = "^\d{2,4}$"
PATT_COUNT_DAYS = "^\d{2,4}$"
PATT_DOG_TYPE = "^ЯЯ"
PATT_DOG_NAME = "^договор займа"
# PATT_DOG_DATE = "^[0-9]{1,2}[.-][0-9]{2}[.-]20[0-9]{2}|^20[0-9]{2}[.-][0-9]{1,2}[0-9]{1,2}"
PATT_DOG_DATE = "[0-9]{1,2}[.-][0-9]{2}[.-]20[0-9]{2}|^20[0-9]{2}[.-][0-9]{1,2}[0-9]{1,2}"
PATT_DOG_NUMBER = "(?:ON)?[A-Za-zА-Яа-я0-9]{1,2}[0-9]{4}20[12][0-9][0-9]{2}|20[0-9]{2}[0-9]{8}"
# PATT_DOG_NUMBER = "^(?:ON)?20[0-9]{2}[0-9]{2}[0-9]{2}[0-9]{4}$|^(?:ON)?[a-zA-Zа-яА-Я0-9]{1,2}[0-9]{6}[0-9]{4}\s*$"
PATT_DOG_PLAT = "Обороты за "
PATT_PAYMENT_DOCUMENT = "Расходный кассовый ордер|Приходный кассовый ордер"
PATT_DOC_PERIOD = "^[0-9]{2}[.][0-9]{2}[.]20[0-9]{2}$"
LEN_DOG_NUMBER = 11

SQL_CONNECT = {
    "dsn": os.environ.get("SQL_DSN", "sqlserverdatasource"),
    "server":os.environ.get("SQL_SERVER", "10.10.10.1"),
    "port": "1433",
    "database": os.environ.get("SQL_DB", "ArchiCreditW"),
    "user": os.environ.get("SQL_USER", "sa"),
    "password": os.environ.get("SQL_PASSWORD", "11111"),
}

LOGGING = {
    "version": 1,
    "disable_existing_loggers": False,
    "filters": {
        "CustomFilter": {
            "()": CustomFilter,
        }
    },
    "formatters": {
        "default": {
            "datefmt": "%Y-%m-%d %H:%M:%S",
            "format": "[%(levelname)s %(asctime)s] %(name)s:%(module)s:%(lineno)d  %(message)s",
        },
        "simple": {
            "()": CustomFormatter,
            "datefmt": "%d-%m-%Y %H:%M:%S",
            "format": "%(asctime)s [%(levelname)s] - {}%(message)s{}",
        },
        "json": {
            "()": "pythonjsonlogger.jsonlogger.JsonFormatter",
            "datefmt": "%Y-%m-%d %H:%M:%S",
            "format": """
                    levelno: %(levelno)s
                    levelname: %(levelname)s
                    asctime: %(asctime)s
                    name: %(name)s
                    module: %(module)s
                    lineno: %(lineno)d
                    message: %(message)s
                    created: %(created)f
                    filename: %(filename)s
                    funcName: %(funcName)s
                    msec: %(msecs)d
                    pathname: %(pathname)s
                    process: %(process)d
                    processName: %(processName)s
                    relativeCreated: %(relativeCreated)d
                    thread: %(thread)d
                    threadName: %(threadName)s
                    exc_info: %(exc_info)s
                """,
            "datefmt": "%Y-%m-%d %H:%M:%S",
        },
    },
    "handlers": {
        "logfile": {
            "formatter": "default",
            "level": "DEBUG",
            "class": "logging.handlers.RotatingFileHandler",
            "encoding": "utf-8",
            "filename": INFO_LOG_FILENAME,
            "maxBytes": 100 * 2**10,
            "backupCount": 2,
        },
        "console": {
            "formatter": "simple",
            "level": "DEBUG",
            "class": "logging.StreamHandler",
            "stream": "ext://sys.stdout",
        },
        "debug": {
            "formatter": "json",
            "level": "DEBUG",
            "class": "logging.handlers.RotatingFileHandler",
            "encoding": "utf-8",
            "filename": DEBUG_LOG_FILENAME,
            "maxBytes": 100 * 2**10,
            "backupCount": 2,
            "delay": True,
        },
        "error": {
            "formatter": "json",
            "level": "WARNING",
            "class": "logging.handlers.RotatingFileHandler",
            "encoding": "utf-8",
            "filename": ERROR_LOG_FILENAME,
            "maxBytes": 100 * 2**10,
            "backupCount": 2,
            "delay": True,
        },
    },
    "loggers": {
        "debug": {
            "level": "DEBUG",
            "handlers": ["console", "debug"],
        },
    },
    "root": {
        "level": "INFO",
        "handlers": [
            "console",
            "logfile",
            "error",
        ],
    },
}

logging.config.dictConfig(LOGGING)
