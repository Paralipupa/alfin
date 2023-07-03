import re, random
from collections import OrderedDict
from datetime import datetime, timedelta
from module.excel_importer import ExcelImporter
from module.excel_exporter import ExcelExporter
from module.helpers import (
    to_date,
    get_order_number,
    get_long_date,
    get_value_without_pattern,
    get_type_pdn,
    get_summa_saldo_end,
    get_attributes,
    get_columns_head,
)
from module.settings import *
from module.data import *

logger = logging.getLogger(__name__)


class Report:
    def __init__(self, filename: str):
        self.order_type = "Основной договор"
        self.name = str(filename)
        self.current_client_key: str = None
        self.clients: OrderedDict = OrderedDict()  # клиенты
        self.suf = (
            "main"
            if self.name.find("58") != -1 or self.name.find("59") != -1
            else "proc"
        )
        self.parser = ExcelImporter(self.name)
        self.current_client_name = ""
        self.current_dogovor_number = ""
        self.current_dogovor_type = "Основной договор"
        self.report_date =  datetime.strptime("31.03.2023","%d.%m.%Y").date() #  datetime.now().date().replace(day=1) - timedelta(days=1)
        self.reference = {}  # ссылки на документы в рамках одного документа
        self.wa = {}
        self.kategoria = {}
        self.warnings = []
        self.fields = {}
        self.tarifs = [(0, "noname"), (1, "Постоянный"), (2, "Старт")]
        self.discounts = (2, 10, 31, 33, 42, 45, 47, 44, 46, 48)

    def __is_read(self) -> bool:
        if self.read():
            self.set_columns()
            return not (
                self.fields.get("FLD_NAME", -1) == -1
                or self.fields.get("FLD_NUMBER", -1) == -1
            )
        return False

    def get_parser(self):
        if self.__is_read():
            for index, self.record in enumerate(self.parser.records):
                self.__record_order_type()
                self.__record_client()
                if self.__get_current_client():
                    self.__record_order_name()
                    self.__record_order_number(index)
                    self.__record_order_date()
                    self.__record_order_period()
                    self.__record_order_payments()
                    self.__record_order_pdn()
                    self.__record_order_rate()
                    self.__record_order_tarif()
                    self.__set_order_count_days()
        return

    def __record_client(self) -> None:
        names = [x for x in self.record[self.fields.get("FLD_NAME")].split(" ") if x]
        if self.__is_find(PATT_NAME, "FLD_NAME") or (
            all(word[0].isupper() for word in names) and len(names) == 3
        ):
            self.current_client_key = (
                self.record[self.fields.get("FLD_NAME")].replace(" ", "").lower()
            )
            self.clients.setdefault(
                self.current_client_key,
                Client(name=self.record[self.fields.get("FLD_NAME")]),
            )
            self.__set_new_order()
            self.__record_order_summa()

    def __get_current_client(self) -> Client:
        if self.current_client_key:
            return self.clients[self.current_client_key]
        else:
            return None

    def __get_current_order(self, is_cashed: bool = False) -> Order:
        client: Client = self.__get_current_client()
        if client:
            return client.order_cache
        else:
            None

    def __get_current_payment(self) -> Payment:
        order = self.__get_current_order()
        if order:
            return order.payment_cache
        else:
            return None

    def __push_current_order(self) -> None:
        client = self.__get_current_client()
        order = self.__get_current_order()
        order.client = client
        client.orders.append(order)

    def __push_current_payment(self) -> None:
        payment = self.__get_current_payment()
        order = self.__get_current_order()
        order.payments_1c.append(payment)
        self.__clean_payment()

    def __set_new_order(self) -> None:
        client: Client = self.__get_current_client()
        if client:
            client.order_cache = Order()

    def __clean_payment(self) -> None:
        order: Order = self.__get_current_order()
        if order:
            order.payment_cache = Payment()

    def __record_order_name(self):
        if self.__is_find("Договор", "FLD_NUMBER"):
            client: Client = self.__get_current_client()
            if client:
                self.__set_new_order()
                order = self.__get_current_order()
                order.type = self.order_type
                order.name = client.name

    def __record_order_type(self):
        if self.__is_find(PATT_DOG_TYPE, "FLD_NUMBER"):
            self.order_type = self.record[self.fields.get("FLD_NUMBER")]

    # Платежи
    def __record_order_payments(self):
        if self.__is_find(PATT_DOG_PLAT, "FLD_NUMBER"):
            date_payment: str = get_value_without_pattern(
                self.record[self.fields["FLD_NUMBER"]], PATT_DOG_PLAT
            )
            if self.__is_find(PATT_CURRENCY, f"FLD_BEG_DEBET_{self.suf}"):
                self.__add_payment(date_payment, "FLD_BEG_DEBET", "B", "D")

            if self.__is_find(PATT_CURRENCY, f"FLD_TURN_DEBET_{self.suf}"):
                self.__add_payment(date_payment, "FLD_TURN_DEBET", "O", "D")

            if self.__is_find(PATT_CURRENCY, f"FLD_TURN_CREDIT_{self.suf}"):
                self.__add_payment(date_payment, "FLD_TURN_CREDIT", "O", "C")

            if self.__is_find(PATT_CURRENCY, f"FLD_END_DEBET_{self.suf}"):
                self.__add_payment(date_payment, "FLD_END_DEBET", "E", "D")

    def __record_order_date(self):
        order = self.__get_current_order()
        if self.__is_find(PATT_DOG_DATE, "FLD_DATE", "date"):
            order.date = to_date(get_long_date(self.record[self.fields["FLD_DATE"]]))
            order.date_begin = order.date
        self.__set_order_field(PATT_DOG_DATE, "FLD_DATE_END", "date_end")
        self.__set_order_field(
            PATT_COUNT_DAYS, "FLD_COUNT_DAYS_DELAY", "count_days_delay"
        )

    def __record_order_pdn(self):
        self.__set_order_field(PATT_PDN, "FLD_PDN", "pdn", True, value_type="float")

    def __record_order_rate(self):
        self.__set_order_field(PATT_RATE, "FLD_RATE", "rate")
        order = self.__get_current_order()
        if order and order.rate:
            order.rate = float(order.rate)

    def __record_order_tarif(self):
        if self.__is_find(PATT_TARIF, "FLD_TARIF"):
            order = self.__get_current_order()
            order.tarif = Tarif()
            name = self.record[self.fields["FLD_TARIF"]]
            if name:
                t = [x for x in self.tarifs if x[1] == name]
                if not t:
                    self.tarifs.append((len(self.tarifs), name))
                    t = self.tarifs[-1]
                else:
                    t = t[0]
                order.tarif.code = t[0]
                order.tarif.name = t[1]

    # Сумма договора в одной из колонок
    def __record_order_summa(self, is_forced: bool = False):
        if self.suf == "main":
            b: bool = self.__set_order_field(
                PATT_CURRENCY,
                f"FLD_TURN_DEBET_{self.suf}",
                "summa",
                is_forced=is_forced,
                value_type="decimal",
            )
            self.__set_order_field(
                PATT_CURRENCY,
                f"FLD_BEG_DEBET_{self.suf}",
                "summa",
                is_forced=is_forced and not b,
                value_type="decimal",
            )
            self.__set_order_field(
                PATT_CURRENCY,
                "FLD_SUMMA",
                "summa",
                is_forced=is_forced and not b,
                value_type="decimal",
            )
        self.__set_order_field(
            PATT_CURRENCY,
            f"FLD_BEG_DEBET_{self.suf}",
            f"debet_beg_{self.suf}",
            is_forced=is_forced,
            value_type="decimal",
        )
        self.__set_order_field(
            PATT_CURRENCY,
            f"FLD_BEG_CREDIT_{self.suf}",
            f"credit_beg_{self.suf}",
            is_forced=is_forced,
            value_type="decimal",
        )
        self.__set_order_field(
            PATT_CURRENCY,
            f"FLD_TURN_DEBET_{self.suf}",
            f"debet_{self.suf}",
            is_forced=is_forced,
            value_type="decimal",
        )
        self.__set_order_field(
            PATT_CURRENCY,
            f"FLD_TURN_CREDIT_{self.suf}",
            f"credit_{self.suf}",
            is_forced=is_forced,
            value_type="decimal",
        )
        self.__set_order_field(
            PATT_CURRENCY,
            f"FLD_END_DEBET_{self.suf}",
            f"debet_end_{self.suf}",
            is_forced=is_forced,
            value_type="decimal",
        )
        self.__set_order_field(
            PATT_CURRENCY,
            f"FLD_END_CREDIT_{self.suf}",
            f"credit_end_{self.suf}",
            is_forced=is_forced,
            value_type="decimal",
        )
        self.__set_order_field(
            PATT_CURRENCY,
            f"FLD_END_DEBET_pen",
            "debet_penalty",
            is_forced=is_forced,
            value_type="decimal",
        )

    def __record_order_period(self):
        self.__set_order_field(PATT_PERIOD, "FLD_PERIOD", "count_days")
        self.__set_order_field(PATT_PERIOD, "FLD_PERIOD_COMMON", "count_days_common")
        order = self.__get_current_order()
        if order:
            if order.count_days:
                order.count_days = int(order.count_days)
            if order.count_days_common:
                order.count_days_common = int(order.count_days_common)

    # Номер договора
    def __record_order_number(self, index: int):
        if self.__is_find(PATT_DOG_NUMBER, "FLD_NUMBER"):
            order = self.__get_current_order()
            order.number = get_order_number(self.record[self.fields.get("FLD_NUMBER")])
            order.row = index
            self.reference.setdefault(order.number, order)
            self.__record_order_summa(True)
            self.__push_current_order()

    def __set_order_count_days(self, order: Order = None):
        if order is None:
            order = self.__get_current_order()
        if order and order.count_days != 0:
            try:
                if order.date_begin:
                    order.date_end = order.date_begin + timedelta(order.count_days)
                    order.count_days_period = self.__get_count_days_in_period(order)
                    order.count_days_common = self.__get_count_days_common(order)
                    order.count_days_delay = self.__get_count_days_delay(order)
            except Exception as ex:
                logger.exception("__set_order_count_days:")

    def __get_order_date_begin(self) -> datetime.date:
        order = self.__get_current_order()
        return order.date_begin

    def __get_order_date_end(self) -> datetime.date:
        # order = self.__get_current_order()
        # date_begin = order.date
        # if date_begin:
        #     order.date_begin = datetime.strftime(date_begin, "%d.%m.%Y")
        #     date_end = date_begin + timedelta(days=float(order.count_days))
        #     order.date_end = datetime.strftime(date_end, "%d.%m.%Y")
        order = self.__get_current_order()
        return order.date_end

    def __get_order_date_calculate(self, order: Order) -> datetime.date:
        if not order.date_calculate:
            order.date_calculate = self.report_date
        return order.date_calculate

    def __get_first_period(self) -> datetime.date:
        return self.report_date.replace(day=1)

    def __get_count_days_in_period(self, order: Order) -> int:
        n_first = 0
        date_first_day_in_month = self.__get_first_period()
        date_last_day_in_month = self.report_date
        order_date = order.date_begin

        # Дата договора в отчетном месяце
        if (order_date > date_first_day_in_month - timedelta(days=1)) and (
            order_date < date_last_day_in_month + timedelta(days=1)
        ):
            date_first_day_in_month = order_date
            n_first = 1
        if order.date_frozen:
            date_end = order.date_frozen
        else:
            date_end = self.__get_order_date_calculate(order)
        # Дата договора в отчетном месяце или ранняя заморозка
        if (
            (date_end > date_first_day_in_month - timedelta(days=1))
            and (date_end < date_last_day_in_month + timedelta(days=1))
            or (date_end < date_first_day_in_month)
        ):
            date_last_day_in_month = date_end
        num_days = (date_last_day_in_month - date_first_day_in_month).days + 1
        date1 = date2 = order_date
        if order.tarif.code == 10 or order.tarif.name.lower() == "старт":
            date1 += timedelta(days=1)
            date2 += timedelta(days=9)
        elif (
            (order.tarif.code == 31)
            or (order.tarif.code == 33)
            or (order.tarif.code == 42)
            or (order.tarif.code == 45)
            or (order.tarif.code == 47)
        ):
            date1 += timedelta(days=1)
            date2 += timedelta(days=6)
        elif (
            (order.tarif.code == 44)
            or (order.tarif.code == 46)
            or (order.tarif.code == 48)
        ):
            date1 += timedelta(days=16)
            date2 += timedelta(days=6)
        if not (
            (date_first_day_in_month == date1) and (date_first_day_in_month == date2)
        ):
            if (
                date1 < date_first_day_in_month
                and date2 > date_first_day_in_month - timedelta(days=1)
            ):
                num_days -= (date2 - date_first_day_in_month + timedelta(days=1)).days
            elif date1 > date_first_day_in_month - timedelta(
                days=1
            ) and date2 < date_last_day_in_month + timedelta(days=1):
                num_days -= (date2 - date1 + timedelta(days=1)).days
            elif (
                date1 > date_first_day_in_month - timedelta(days=1)
                and date1 < date_last_day_in_month + timedelta(days=1)
                and date2 > date_last_day_in_month
            ):
                num_days -= (date_last_day_in_month - date1 + timedelta(days=1)).days
            elif (
                date1 < date_first_day_in_month
                and date1 < date_last_day_in_month > date_last_day_in_month
            ):
                num_days = 0
        num_days -= n_first
        return num_days if num_days > 0 else 0

    def __get_count_days_common(self, order: Order = None) -> int:
        if order is None:
            order: Order = self.__get_current_order()
        last_day_on_period: datetime.date = self.report_date
        count_days = (last_day_on_period - order.date_begin).days
        if order.tarif.code in self.discounts:
            count_days -= 7
        return count_days if count_days > 0 else 0

    # Количество дней просрочки
    def __get_count_days_delay(self, order: Order) -> int:
        last_day_on_period: datetime.date = self.report_date
        try:
            return (
                (
                    last_day_on_period - order.date_begin - timedelta(order.count_days)
                ).days
                if last_day_on_period > order.date_end
                else 0
            )
        except Exception as ex:
            logger.exception("__get_count_days_delay:")

    def __is_find(
        self,
        pattern: str,
        column_fld_name: str = None,
        order_fld_name: str = None,
        is_forced: bool = False,
    ) -> bool:
        order = self.__get_current_order()
        return (
            self.fields.get(column_fld_name, -1) != -1
            and (
                order_fld_name is None
                or (not getattr(order, order_fld_name) or is_forced)
            )
            and re.search(
                pattern, self.record[self.fields.get(column_fld_name)], re.IGNORECASE
            )
        )

    def __set_order_field(
        self,
        pattern: str,
        column_fld_name: str = None,
        order_fld_name: str = "",
        is_forced: bool = False,
        value_type: str = "",
    ) -> bool:
        if self.__is_find(pattern, column_fld_name, order_fld_name, is_forced):
            order = self.__get_current_order()
            if value_type == "decimal":
                setattr(
                    order,
                    order_fld_name,
                    Decimal(self.record[self.fields[column_fld_name]]),
                )
            elif value_type == "float":
                setattr(
                    order,
                    order_fld_name,
                    float(self.record[self.fields[column_fld_name]]),
                )
            else:
                setattr(
                    order, order_fld_name, self.record[self.fields[column_fld_name]]
                )
            return True
        return False

    def __add_payment(
        self, date_payment: str, fld_name: str, p_type: str, p_category: str
    ):
        if self.__is_find(PATT_CURRENCY, f"{fld_name}_{self.suf}"):
            payment = self.__get_current_payment()
            payment.summa = Decimal(
                self.record[self.fields.get(f"{fld_name}_{self.suf}")]
            )
            payment.date = to_date(get_long_date(date_payment))
            payment.kind = self.suf
            payment.type = p_type
            payment.category = p_category
            self.__push_current_payment()

    # Устанавливаем номера колонок
    def set_columns(self):
        items = get_columns_head(self.suf)
        index = 0
        for record in self.parser.records:
            col = 0
            for cell in record:
                if re.search("Оборотно-сальдовая ведомость по счету", cell):
                    x = re.findall("(?<=г\. - ).+(?= г[\.])", cell)
                    if x:
                        self.report_date = to_date(x[0])

                for item in items:
                    for name in item["name"]:
                        if not self.fields.get(name) and re.search(
                            item["pattern"], cell
                        ):
                            self.fields[name] = col + item["off_col"]
                col += 1
            if self.fields.get("FLD_NAME", -1) != -1:
                return
            index += 1
            if index > 20:
                return

    def read(self) -> bool:
        return self.parser.read()

    def write_to_excel(self, filename: str = "output_full") -> str:
        exel = ExcelExporter("output_excel")
        return exel.write(self)

    def union_all(self, items):
        pattern: re.Pattern = re.compile(
            "_proc|_main|pdn|rate|count_|date|percent"
        )
        if not items:
            return
        order_attrs = [x for x in get_attributes(Order()) if pattern.search(x)]
        order: Order
        client: Client
        for key, client in self.clients.items():
            client_items = [
                x.clients[key]
                for x in items
                if x.clients.get(key) and len(x.clients[key].orders) == 0
            ]
            for order in client.orders:
                order_items = [
                    x.reference[order.number]
                    for x in items
                    if x.reference.get(order.number)
                ]
                order_item: Order
                for order_item in order_items:
                    for attr in order_attrs:
                        if not getattr(order, attr) and getattr(order_item, attr):
                            setattr(order, attr, getattr(order_item, attr))
                    if order_item.tarif.code != 0:
                        order.tarif = order_item.tarif
                    payment: Payment
                    for payment in order_item.payments_1c:
                        order.payments_1c.append(payment)
                        if (
                            order.date_frozen is None
                            and payment.date
                            and payment.summa
                            and payment.type == "O"
                            and payment.category == "C"
                        ):
                            order.date_frozen = payment.date
                if client_items:
                    for client_item in client_items:
                        for attr in order_attrs:
                            if not getattr(order, attr) and getattr(
                                client_item.order_cache, attr
                            ):
                                setattr(
                                    order, attr, getattr(client_item.order_cache, attr)
                                )

    def fill_from_archi(self, data: dict):
        if not data:
            return
        for client in self.clients.values():
            order: Order = Order()
            for order in client.orders:
                if data["order"].get(order.number):
                    order.rate = data["order"][order.number][1]
                    order.tarif.code = data["order"][order.number][6]
                    order.tarif.name = data["order"][order.number][7]
                    order.count_days = data["order"][order.number][2]
                    self.__set_order_count_days(order)
                if data["payment"].get(order.number):
                    for payment in data["payment"][order.number]:
                        order.payments_base.append(payment)

    # средневзвешенная величина
    def set_weighted_average(self):
        item: Client = Client()
        order: Order = Order()
        for item in self.clients.values():
            for order in item.orders:
                period = order.count_days
                summa = order.summa
                tarif = order.tarif.code
                tarif_name = order.tarif.name
                rate = order.rate
                if period and summa and tarif and rate:
                    key = f"{tarif_name}_{rate}"
                    data = self.wa.get(key)
                    period = float(period)
                    if not data:
                        # 46 -Друг
                        self.wa[key] = {
                            "parent": [],
                            "stavka": float(rate),
                            "koef": 240.194
                            if tarif in self.discounts
                            else 365 * float(rate),
                            "period": period - 7 if tarif in self.discounts else period,
                            "summa_free": 0,
                            "summa": 0,
                            "count": 0,
                            "value": {},
                        }
                    self.wa[key]["parent"].append(order)
                    s = self.wa[key]["value"].get(summa)
                    if not s:
                        self.wa[key]["value"][summa] = 1
                    else:
                        self.wa[key]["value"][summa] += 1
                    self.wa[key]["summa"] += float(summa) * self.wa[key]["koef"]
                    self.wa[key]["summa_free"] += float(summa)
                    self.wa[key]["count"] += 1
                else:
                    if summa:
                        self.warnings.append(
                            f"ср.взвеш: {item.name} {order.number}  {summa} period:{period} tarif:{tarif} proc:{rate}"
                        )
        summa = 0
        summa_free = 0
        for key, item in self.wa.items():
            summa += item["summa"]
            summa_free += item["summa_free"]
        self.wa["summa_free"] = summa_free
        self.wa["summa"] = summa
        self.wa["summa_wa"] = summa / summa_free if summa_free != 0 else 1

    # категории потребительских займов
    def set_kategoria(self):
        def get_kategoria() -> dict:
            d = {}
            data = [
                (1, "30"),
                (2, "40"),
                (3, "50"),
                (4, "60"),
                (5, "70"),
                (6, "80"),
                (7, "99"),
                (0, ""),
            ]
            for x in data:
                d[str(x[0])] = {
                    "title": x[1],
                    "count4": 0,
                    "count6": 0,
                    "summa5": 0,
                    "summa3": 0,
                    "summa6": 0,
                    "items": [],
                }
            return d

        kategoria = get_kategoria()
        reserves = {}
        client: Client = Client()
        order: Order = Order()
        random.seed()
        for client in self.clients.values():
            pdn = 0  ## random.random()
            for order in client.orders:
                if order.pdn:
                    if order.pdn > 1:
                        pdn = order.pdn / 100
                order.pdn = round(pdn, 2)
            for order in client.orders:
                t = get_type_pdn(order.summa, order.pdn)
                summa = get_summa_saldo_end(order)
                kategoria[t]["count4"] += 1
                kategoria[t]["summa5"] += order.summa
                kategoria[t]["summa3"] += summa
                if order.count_days_delay > 90 and summa > 0:
                    kategoria[t]["count6"] += 1
                    kategoria[t]["summa6"] += summa
                item = {"name": client.name, "parent": order}
                kategoria[t]["items"].append(item)
                order.percent = self.__get_rezerv_percent(order.count_days_delay)
                reserves.setdefault(str(order.percent), Reserve())
                reserves[str(order.percent)].percent = order.percent
                reserves[str(order.percent)].count += 1
                reserves[str(order.percent)].items.append(item)
        reserves = sorted(reserves.items())
        for item in reserves:
            item[1].items = sorted(item[1].items, key=lambda x: x["name"])
        self.reserve = reserves
        self.kategoria = kategoria

    def get_numbers(self):
        return [
            f"0{x.split()}" if len(x.split()) == 11 else x
            for x in self.reference.keys()
        ]

    def __get_rezerv_percent(self, count: int) -> float:
        if count <= 7:
            return 0
        elif count <= 30:
            return 3 / 100
        elif count <= 60:
            return 10 / 100
        elif count <= 90:
            return 20 / 100
        elif count <= 120:
            return 40 / 100
        elif count <= 180:
            return 50 / 100
        elif count <= 270:
            return 65 / 100
        elif count <= 360:
            return 80 / 100
        else:
            return 99 / 100
