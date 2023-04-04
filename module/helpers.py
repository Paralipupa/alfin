import datetime
import logging

logger = logging.getLogger(__name__)


def timing(start_message: str = "", end_message: str = "Завершено"):
    def wrap(f):
        def inner(*args, **kwargs):
            if start_message:
                logger.info(start_message)
            time1 = datetime.datetime.now()
            ret = f(*args, **kwargs)
            time2 = datetime.datetime.now()
            if end_message:
                logger.info("{0} ({1})".format(end_message, (time2 - time1)))
            return ret

        return inner

    return wrap

def last_day_of_month(any_day):
    next_month = any_day.replace(day=28) + datetime.timedelta(
        days=4
    ) 
    d = next_month - datetime.timedelta(days=next_month.day)
    return d.date()


def to_date(x: str):
    months = [
        ("Январь", "January"),
        ("Февраль", "February"),
        ("Март", "March"),
        ("Апрель", "April"),
        ("Май", "May"),
        ("Июнь", "June"),
        ("Июль", "July"),
        ("Август", "August"),
        ("Сентябрь", "September"),
        ("Октябрь", "October"),
        ("Ноябрь", "November"),
        ("Декабрь", "December"),
    ]
    for mon in months:
        x = x.replace(mon[0], mon[1])
    try:
        d = datetime.datetime.strptime(x, "%B %Y")
        return last_day_of_month(d)
    except:
        pass
    d = get_date(x)
    if d:
        return d
    return x

def get_date(date_str: str) -> datetime.datetime:
    patts = [
        "%d.%m.%Y",
        "%d.%m.%y",
        "%d.%m.%Y %H:%M:%S",
        "%d.%m.%y %H:%M:%S",
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%d-%m-%y",
        "%d/%m/%y",
        "%B %Y",
    ]
    d = None
    for p in patts:
        try:
            d = datetime.datetime.strptime(date_str, p)
            return d.date()
        except:
            pass
    return None
