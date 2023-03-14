import datetime
import logging

logger = logging.getLogger(__name__)

def timing(start_message:str="", end_message:str = "Завершено"):
    def wrap(f):
        def inner(*args, **kwargs):
            if start_message: logger.info(start_message)
            time1 = datetime.datetime.now()
            ret = f(*args, **kwargs)
            time2 = datetime.datetime.now()
            if end_message: logger.info("{0} ({1})".format(end_message, (time2 - time1)))
            return ret
        return inner
    return wrap
