import hashlib
from typing import List
from dataclasses import dataclass, field
from datetime import datetime, date
from decimal import Decimal


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
    summa: Decimal = 0


@dataclass
class Order:
    type: str = None
    name: str = None
    number: str = None
    date: date = None
    date_begin: date = None
    date_end: date = None
    summa: Decimal = 0
    rate: float = 0
    percent: float = 0
    tarif: Tarif = Tarif()
    count_days: int = 0
    count_days_common: int = 0
    count_days_period: int = 0
    count_days_delay: int = 0
    pdn: float = 0.3
    payments_1c: List[Payment] = field(default_factory=list)
    payments_base: List[tuple] = field(default_factory=list)
    payment_cache: Payment = Payment()
    date_frozen: date = None
    date_calculate: date = None
    row: int = 0
    is_cashed: bool = False
    link: dict = field(default_factory=dict)


@dataclass
class Client:
    name: str = None
    orders: List[Order] = field(default_factory=list)
    order_cache: Order = Order()
    link: dict = field(default_factory=dict)


@dataclass
class Reserve:
    percent: float = 0
    summa_main: Decimal = 0
    summa_percent: Decimal = 0
    count: int = 0
    items: list = field(default_factory=list)


def get_kategoria():
    # +1!YtqY6dX
    return {
        "1": {
            "title": "30",
            "count4": 0,
            "count6": 0,
            "summa5": 0,
            "summa3": 0,
            "summa6": 0,
            "items": [],
        },
        "2": {
            "title": "40",
            "count4": 0,
            "count6": 0,
            "summa5": 0,
            "summa3": 0,
            "summa6": 0,
            "items": [],
        },
        "3": {
            "title": "50",
            "count4": 0,
            "count6": 0,
            "summa5": 0,
            "summa3": 0,
            "summa6": 0,
            "items": [],
        },
        "4": {
            "title": "60",
            "count4": 0,
            "count6": 0,
            "summa5": 0,
            "summa3": 0,
            "summa6": 0,
            "items": [],
        },
        "5": {
            "title": "70",
            "count4": 0,
            "count6": 0,
            "summa5": 0,
            "summa3": 0,
            "summa6": 0,
            "items": [],
        },
        "6": {
            "title": "80",
            "count4": 0,
            "count6": 0,
            "summa5": 0,
            "summa3": 0,
            "summa6": 0,
            "items": [],
        },
        "7": {
            "title": "99",
            "count4": 0,
            "count6": 0,
            "summa5": 0,
            "summa3": 0,
            "summa6": 0,
            "items": [],
        },
        "0": {
            "title": "",
            "count4": 0,
            "count6": 0,
            "summa5": 0,
            "summa3": 0,
            "summa6": 0,
            "items": [],
        },
    }
