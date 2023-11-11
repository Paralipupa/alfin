import hashlib, re
from typing import List
from dataclasses import dataclass, field
from datetime import datetime, date
from decimal import Decimal
from .settings import PATT_TIME_IN_DOCUMENT


def _hashit(s):
    return hashlib.sha1(s).hexdigest()


class HashMixin(object):
    pass

    def get_hash(self):
        attrs = [a for a in dir(self) if not a.startswith("__")]
        hashed_fields = [str(getattr(self, field)) for field in attrs]
        data = ",".join(hashed_fields)
        return _hashit(data.encode("utf-8"))

    hash = property(get_hash)


@dataclass
class Tarif:
    code: int = 0
    name: str = ""


@dataclass
class Payment(HashMixin):
    kind: str = None
    type: str = None  # Сальдо на начала, Обороты, Сальдо на конец
    category: str = None  # Debet,  Credit
    date: date = None
    account_debet: str = ""  #  Счет
    account_credit: str = ""  #  Счет
    summa: Decimal = 0

    def get_account(self, a: str) -> str:
        if a == "66.03":
            return "42316"
        if a == "58.03":
            return "48801"
        if a == "76.06":
            return "48809"
        if a == "76.09":
            return "48809"
        if a == "71.01":
            return "60308"
        if a == "94":
            return "60323"
        if a == "98.02":
            return "10614"
        if a == "75.01":
            return "60330"
        if a == "70":
            return "60305"
        return a


@dataclass
class Order:
    client: dict = None
    type: str = None
    name: str = None
    number: str = None
    date_order: datetime.date = None
    date_begin: datetime.date = None
    date_end: datetime.date = None
    summa: Decimal = 0
    credit_main: Decimal = 0
    rate: float = 0
    percent: float = 0
    tarif: Tarif = Tarif()
    count_days: int = 0
    count_days_common: int = 0
    count_days_period: int = 0
    count_days_delay: int = 0
    pdn: float = 0
    payments_1c: List[Payment] = field(default_factory=list)
    payments_base: List[tuple] = field(default_factory=list)
    payment_cache: Payment = Payment()
    date_frozen: datetime.date = None
    date_calculate: datetime.date = None
    row: int = 0
    is_cashed: bool = False
    debet_beg_main: Decimal = 0
    credit_beg_main: Decimal = 0  #
    debet_main: Decimal = 0  # начислено по основному долгу
    credit_main: Decimal = 0  # оплачено по основному долгу
    debet_end_main: Decimal = 0
    credit_end_main: Decimal = 0  # кредит по основному долгу
    debet_beg_proc: Decimal = 0
    credit_beg_proc: Decimal = 0  #
    debet_proc: Decimal = 0  # начислено по процентам
    credit_proc: Decimal = 0  # оплачено по процентам
    debet_end_proc: Decimal = 0
    debet_end_proc_58: Decimal = 0  # данные из 58рез1 (сумма начисл. процентов)
    credit_end_proc: Decimal = 0  # кредит по остатку
    debet_penalty: Decimal = 0
    summa_percent_period: Decimal = 0
    summa_percent_all: Decimal = 0
    summa_reserve_main: Decimal = 0
    summa_reserve_proc: Decimal = 0
    summa_reserve_main_58: Decimal = 0  # данные из 58рез1 (резерв по основному долгу)
    summa_reserve_main_58_pdn: Decimal = (
        0  # данные из 58рез1 (резерв по основному долгу ПДН)
    )
    summa_reserve_proc_58: Decimal = 0  # данные из 58рез1 (резерв по процентам)
    summa_reserve_proc_58_pdn: Decimal = 0  # данные из 58рез1 (резерв по процентам ПДН)
    summa_payment: Decimal = 0
    calc_reserve_main: Decimal = 0
    calc_reserve_proc: Decimal = 0
    link: dict = field(default_factory=dict)

    def get_date(self):
        try:
            date_order_str = "{}.{}.{}".format(
                self.number[-6:-2], self.number[-8:-6], self.number[-10:-8]
            )
            date_order = datetime.strptime(date_order_str, "%Y.%m.%d")
            return date_order
        except:
            return None


@dataclass
class Document:
    text: str = ""
    number: str = ""
    date_period: date = None
    summa: float = 0
    code: str = ""
    basis: str = ""
    order: Order = None
    client = None
    is_print: bool = False

    def __init__(self, text: str, debet: str = "", credit: str = ""):
        self.text = text
        self.is_print = False
        if bool(debet):
            self.code = "1"
        elif bool(credit):
            self.code = "2"
        result = re.search("(?<=BZ )[а-яА-Яa-zA-Z0-9]+|(?<=BZ)[а-яА-Яa-zA-Z0-9]+", text)
        if result:
            self.number = result.group(0).strip()
        pattern = f"{PATT_TIME_IN_DOCUMENT}\s.+(?=адрес)"
        pattern = f"|{PATT_TIME_IN_DOCUMENT}\s.+(?=\()"
        pattern = f"|{PATT_TIME_IN_DOCUMENT}\s.+"
        result = re.search(pattern, text)
        if result:
            self.basis = re.sub(
                f"{PATT_TIME_IN_DOCUMENT}|<...>", "", result.group(0).strip()
            )
        return


@dataclass
class Client:
    name: str = None
    account: str = ""
    orders: List[Order] = field(default_factory=list)
    order_cache: Order = Order()
    documents: List[Document] = field(default_factory=list)
    passport_number: str = ""
    link: dict = field(default_factory=dict)


@dataclass
class Reserve:
    percent: float = 0
    count: int = 0
    items: list = field(default_factory=list)

    # +1!YtqY6dX
