import datetime, os, pathlib, json
import logging
from decimal import Decimal
from typing import Any
from module.settings import LEN_DOG_NUMBER, ENCONING
from module.data import Order
from .settings import PATT_NAME, PATT_DOG_NUMBER

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
    next_month = any_day.replace(day=28) + datetime.timedelta(days=4)
    d = next_month - datetime.timedelta(days=next_month.day)
    return d.date()


def to_date(x: str) -> datetime.date:
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
    return None


def get_date(date_str: str) -> datetime.date:
    patts = [
        "%d.%m.%Y",
        "%d.%m.%y",
        "%d.%m.%Y %H:%M:%S",
        "%d.%m.%y %H:%M:%S",
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
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


def get_order_number(number: str) -> str:
    return (
        f"0{number.strip()}"
        if len(number.strip()) == LEN_DOG_NUMBER
        else number.strip()
    )


def get_long_date(text: str) -> str:
    return text + (" 0:00:00" if text.find(":") == -1 else "")


def get_value_without_pattern(text: str, pattern: str) -> str:
    return text.replace(pattern, "").strip()


def write(filename: str = "output", docs: list = []):
    os.makedirs("output", exist_ok=True)
    with open(
        pathlib.Path("output", f"{filename}.json"),
        mode="w",
        encoding=ENCONING,
    ) as file:
        jstr = json.dumps(docs, indent=4, ensure_ascii=False)
        file.write(jstr)


def get_type_pdn(summa: Decimal, pdn: float) -> str:
    if summa >= 10000:
        if pdn <= 0.3:
            t = "1"
        elif pdn <= 0.4:
            t = "2"
        elif pdn <= 0.5:
            t = "3"
        elif pdn <= 0.6:
            t = "4"
        elif pdn <= 0.7:
            t = "5"
        elif pdn <= 0.8:
            t = "6"
        else:
            t = "7"
    else:
        t = "0"
    return t


def get_summa_saldo_end(order: Order):
    return sum(
        [x.summa for x in order.payments_1c if x.type == "E" and x.category == "D"]
    )


def get_summa_turn_main(order: Order):
    return sum(
        [
            x.summa
            for x in order.payments_1c
            if x.type == "O" and x.category == "D" and x.kind == "main"
        ]
    )


def get_summa_turn_percent(order: Order):
    return sum(
        [
            x.summa
            for x in order.payments_1c
            if x.type == "O" and x.category == "D" and x.kind == "proc"
        ]
    )


def get_value_attr(attr_value: str, type_attr: str) -> Any:
    if type_attr == "float":
        value = float(attr_value)
    elif type_attr == "int":
        value = int(attr_value)
    elif type_attr == "date":
        value = to_date(attr_value)
    elif type_attr == "str":
        value = str(attr_value)
    return value


def get_max_margin_rate(ddate: datetime.datetime.date) -> float:
    if ddate < datetime.datetime.strptime("28.01.2019", "%d.%m.%Y").date():
        return 3
    elif ddate < datetime.datetime.strptime("01.07.2019", "%d.%m.%Y").date():
        return 2.5
    elif ddate < datetime.datetime.strptime("01.01.2020", "%d.%m.%Y").date():
        return 2
    elif ddate < datetime.datetime.strptime("01.07.2023", "%d.%m.%Y").date():
        return 1.5
    else:
        return 1.3


def get_attributes(obj) -> list:
    return [x for x in dir(obj) if not x.startswith("__")]


def get_columns_head(suf: str):
    return [
        {"name": ["FLD_NAME"], "pattern": f"{PATT_NAME}|^Контрагент$", "off_col": 0},
        {"name": ["FLD_NUMBER", "FLD_DATE"], "pattern": f"{PATT_DOG_NUMBER}|^Контрагент$", "off_col": 0},
        {"name": ["FLD_DATE"], "pattern": "^Дата выдачи", "off_col": 0},
        {
            "name": ["FLD_NUMBER"],
            "pattern": "^№ заявки$|^Договор$|^№ договора$",
            "off_col": 0,
        },
        {
            "name": [f"FLD_BEG_DEBET_{suf}"],
            "pattern": "^Сальдо на начало периода$",
            "off_col": 0,
        },
        {
            "name": [f"FLD_BEG_CREDIT_{suf}"],
            "pattern": "^Сальдо на начало периода$",
            "off_col": 1,
        },
        {
            "name": [f"FLD_TURN_DEBET_{suf}", "FLD_SUMMA"],
            "pattern": "^Обороты за период$",
            "off_col": 0,
        },
        {
            "name": [f"FLD_TURN_CREDIT_{suf}"],
            "pattern": "^Обороты за период$",
            "off_col": 1,
        },
        {
            "name": [f"FLD_END_DEBET_{suf}"],
            "pattern": "^Сальдо на конец периода$",
            "off_col": 0,
        },
        {
            "name": [f"FLD_BEG_DEBET_{suf}"],
            "pattern": "^Дебет$",
            "off_col": 1,
        },
        {
            "name": [f"FLD_END_DEBET_{suf}"],
            "pattern": "^Дебет$",
            "off_col": 1,
        },
        {
            "name": [f"FLD_END_CREDIT_{suf}"],
            "pattern": "^Сальдо на конец периода$",
            "off_col": 1,
        },
        # {
        #     "name": [f"FLD_END_DEBET_main"],
        #     "pattern": "^Основной долг$",
        #     "off_col": 0,
        # },
        # {
        #     "name": [f"FLD_END_DEBET_proc"],
        #     "pattern": "^Долг по процентам$",
        #     "off_col": 0,
        # },
        {
            "name": [f"FLD_END_DEBET_pen"],
            "pattern": "^Долг по единовременным штрафам$",
            "off_col": 0,
        },
        {
            "name": ["FLD_PERIOD"],
            "pattern": "^Общий срок займа$|^Общая сумма долга по процентам$",
            "off_col": 0,
        },
        {
            "name": ["FLD_PERIOD_COMMON"],
            "pattern": "^кол-во дней для расчета проц\.$|^Общий срок займа$",
            "off_col": 0,
        },
        {
            "name": ["FLD_RATE"],
            "pattern": "^Общая сумма долга по процентам$",
            "off_col": 1,
        },
        {"name": ["FLD_RATE"], "pattern": "^Процентная ставка", "off_col": 0},
        {
            "name": ["FLD_TARIF"],
            "pattern": "^Общая сумма долга по процентам$|^Наименование продукта$",
            "off_col": 0,
        },
        {
            "name": ["FLD_COUNT_DAYS"],
            "pattern": "^кол-во дней начисления процента$",
            "off_col": 0,
        },
        {
            "name": ["FLD_COUNT_DAYS_DELAY"],
            "pattern": "^кол-во дней просрочки до отчетного периода$",
            "off_col": 0,
        },
        {"name": ["FLD_PDN"], "pattern": "^Показатель долговой|ПДН", "off_col": 0},
        {
            "name": [f"FLD_END_DEBET_proc_58"],
            "pattern": "^сумма начисл\. процентов$",
            "off_col": 0,
        },
        {"name": ["FLD_DATE_BEGIN"], "pattern": "^Дата пролонгации", "off_col": 0},
        {"name": ["FLD_DATE_FROZEN"], "pattern": "^дата заморозки$", "off_col": 0}, 
        {
            "name": ["FLD_SUMMA"],
            "pattern": "Сумма займа|^Выданная сумма займа$",
            "off_col": 0,
        },
        {
            "name": ["FLD_SUMMA_PAYMENT"],
            "pattern": "^Общая сумма основного долга$", # из 58,03рез1
            "off_col": 0,
        },
        {
            "name": ["FLD_SUMMA_RESERVE_MAIN"],
            "pattern": "^резерв по основному долгу$", # из 58,03рез1
            "off_col": 0,
        },
        {
            "name": ["FLD_SUMMA_RESERVE_MAIN_PDN"],
            "pattern": "^резерв по основному долгу$", # из 58,03рез1
            "off_col": 1,
        },
        {
            "name": ["FLD_SUMMA_RESERVE_PROC"],
            "pattern": "^резерв по процентам$", # из 58,03рез1
            "off_col": 0,
        },
        {
            "name": ["FLD_SUMMA_RESERVE_PROC_PDN"],
            "pattern": "^резерв по процентам$", # из 58,03рез1
            "off_col": 1,
        },
    ]
