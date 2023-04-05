import re
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
    get_summa_turn_main,
    get_summa_turn_percent,
)
from module.data import get_kategoria
from module.serializer import serializer, deserializer
from module.settings import *
from module.data import *

logger = logging.getLogger(__name__)


class Report:
    def __init__(self, filename: str):
        self.order_type = "Основной договор"
        self.name = str(filename)
        self.clients = list()  # клиенты
        self.suf = "main" if self.name.find("58") != -1 else "proc"
        self.parser = ExcelImporter(self.name)
        self.current_client_name = ""
        self.current_dogovor_number = ""
        self.current_dogovor_type = "Основной договор"
        self.report_date = datetime.now().date().replace(day=1) - timedelta(days=1)
        self.reference = {}  # ссылки на документы в рамках одного документа
        self.wa = {}
        self.kategoria = {}
        self.warnings = []
        self.fields = {}

    def __is_read(self) -> bool:
        self.read()
        self.set_columns()
        return not (
            self.fields.get("FLD_NAME", -1) == -1
            or self.fields.get("FLD_NUMBER", -1) == -1
        )

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

    def __record_client(self) -> None:
        if self.__is_find(PATT_NAME, "FLD_NAME"):
            self.clients.append(Client(name=self.record[self.fields.get("FLD_NAME")]))

    def __get_current_client(self) -> Client:
        if len(self.clients) != 0:
            return self.clients[-1]
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
            value: str = get_value_without_pattern(
                self.record[self.fields["FLD_NUMBER"]], PATT_DOG_PLAT
            )
            if self.__is_find(PATT_CURRENCY, f"FLD_BEG_DEBET_{self.suf}"):
                self.__add_payment(value, "FLD_BEG_DEBET", "B", "D")

            if self.__is_find(PATT_CURRENCY, f"FLD_TURN_DEBET_{self.suf}"):
                self.__add_payment(value, "FLD_TURN_DEBET", "O", "D")

            if self.__is_find(PATT_CURRENCY, f"FLD_TURN_CREDIT_{self.suf}"):
                self.__add_payment(value, "FLD_TURN_CREDIT", "O", "C")

            if self.__is_find(PATT_CURRENCY, f"FLD_END_DEBET_{self.suf}"):
                self.__add_payment(value, "FLD_END_DEBET", "E", "D")

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
        self.__set_order_field(PATT_PDN, "FLD_PDN", "pdn")

    def __record_order_rate(self):
        self.__set_order_field(PATT_RATE, "FLD_RATE", "rate")

    def __record_order_tarif(self):
        self.__set_order_field(PATT_TARIF, "FLD_TARIF", "tarif")

    # Сумма договора в одной из колонок
    def __record_order_summa(self, is_forced: bool = False):
        if self.suf == "main":
            order = self.__get_current_order()
            if self.__is_find(PATT_CURRENCY, "FLD_SUMMA", "summa", is_forced=is_forced):
                order.summa = Decimal(self.record[self.fields["FLD_SUMMA"]])
            if self.__is_find(PATT_CURRENCY, "FLD_BEG_DEBET", "summa"):
                order.summa = Decimal(self.record[self.fields["FLD_BEG_DEBET"]])
            if self.__is_find(PATT_CURRENCY, "FLD_TURN_DEBET", "summa"):
                order.summa = Decimal(self.record[self.fields["FLD_TURN_DEBET"]])
            if self.__is_find(PATT_CURRENCY, "FLD_END_DEBET", "summa"):
                order.summa = Decimal(self.record[self.fields["FLD_END_DEBET"]])

    def __record_order_period(self):
        self.__set_order_field(PATT_PERIOD, "FLD_PERIOD", "period")
        self.__set_order_field(PATT_PERIOD, "FLD_PERIOD_COMMON", "period_common")

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
        npoid = order.tarif
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
        if npoid == 10:
            date1 += timedelta(days=1)
            date2 += timedelta(days=9)
        elif (
            (npoid == 31)
            or (npoid == 33)
            or (npoid == 42)
            or (npoid == 45)
            or (npoid == 47)
        ):
            date1 += timedelta(days=1)
            date2 += timedelta(days=6)
        elif (npoid == 44) or (npoid == 46) or (npoid == 48):
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
        npoid = order.tarif
        count_days = (last_day_on_period - order.date_begin).days
        if (
            (npoid == 10)
            or (npoid == 31)
            or (npoid == 33)
            or (npoid == 42)
            or (npoid == 45)
            or (npoid == 47)
            or (npoid == 44)
            or (npoid == 46)
            or (npoid == 48)
        ):
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
    ) -> None:
        if self.__is_find(pattern, column_fld_name, order_fld_name, is_forced):
            order = self.__get_current_order()
            setattr(order, order_fld_name, self.record[self.fields[column_fld_name]])

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
        items = [
            {"name": ["FLD_NAME"], "pattern": "^Счет$|^ФИО|^Контрагент$", "off_col": 0},
            {
                "name": [f"FLD_BEG_DEBET_{self.suf}"],
                "pattern": "^Сальдо на начало периода$",
                "off_col": 0,
            },
            {
                "name": [f"FLD_TURN_DEBET_{self.suf}", "FLD_SUMMA"],
                "pattern": "^Обороты за период$",
                "off_col": 0,
            },
            {
                "name": [f"FLD_TURN_CREDIT_{self.suf}"],
                "pattern": "^Обороты за период$",
                "off_col": 1,
            },
            {
                "name": [f"FLD_END_DEBET_{self.suf}"],
                "pattern": "^Сальдо на конец периода$",
                "off_col": 0,
            },
            {
                "name": ["FLD_PERIOD"],
                "pattern": "^Первоначальный срок займа$|^Общая сумма долга по процентам$",
                "off_col": 0,
            },
            {
                "name": ["FLD_PERIOD_COMMON"],
                "pattern": "^кол-во дней для расчета проц\.$",
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
                "name": ["FLD_END_DEBET_{self.suf}"],
                "pattern": "^сумма начисл\. процентов$",
                "off_col": 0,
            },
            {"name": ["FLD_NUMBER", "FLD_DATE"], "pattern": "^Счет$", "off_col": 0},
            {
                "name": ["FLD_NUMBER"],
                "pattern": "^№ заявки$|^Договор$|^№ договора$",
                "off_col": 0,
            },
            {"name": ["FLD_DATE"], "pattern": "^Дата выдачи", "off_col": 0},
            {
                "name": ["FLD_SUMMA"],
                "pattern": "Сумма займа|^Выданная сумма займа$",
                "off_col": 0,
            },
        ]
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

    def read(self):
        self.parser.read()

    def write_to_excel(self, filename: str = "output_full") -> str:
        exel = ExcelExporter("output_excel")
        return exel.write(self)

    def union_all(self, items):
        if not items:
            return
        payment: Payment = Payment()
        order: Order = Order()
        order_sub: Order = Order()
        for number, order in self.reference.items():
            for order_sub in [
                x.reference[number] for x in items if x.reference.get(number)
            ]:
                for payment in order_sub.payments_1c:
                    order.payments_1c.append(payment)
                    if (
                        order.date_frozen is None
                        and payment.date
                        and payment.summa
                        and payment.type == "O"
                        and payment.category == "C"
                    ):
                        order.date_frozen = payment.date
        # write(self.clients)

    def fill_from_archi(self, data: dict):
        if not data:
            return
        for client in self.clients:
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
        for item in self.clients:
            for order in item.orders:
                period = order.count_days
                summa = order.summa
                tarif = order.tarif.code
                rate = order.rate
                if period and summa and tarif and rate:
                    key = f"{tarif}_{rate}"
                    data = self.wa.get(key)
                    period = float(period)
                    if not data:
                        # 46 -Друг
                        self.wa[key] = {
                            "parent": [],
                            "stavka": float(rate),
                            "koef": 240.194
                            if tarif == 46 or tarif == 48
                            else 365 * float(rate),
                            "period": period - 7
                            if tarif == 46 or tarif == 48
                            else period,
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
        kategoria = get_kategoria()
        reserves = {}
        client: Client = Client()
        order: Order = Order()
        for client in self.clients:
            pdn = 0.3
            for order in client.orders:
                pdn = order.pdn if order.pdn else pdn
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

                reserves.setdefault(str(order.percent), Reserve())
                reserves[str(order.percent)].summa_main += get_summa_turn_main(order)
                reserves[str(order.percent)].summa_percent += get_summa_turn_percent(
                    order
                )
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
