import re
import random
from collections import OrderedDict
from datetime import datetime, timedelta
from dateutil import parser
from typing import List
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
from .helpers import get_max_margin_rate
from .error_report import ErrorReport


logger = logging.getLogger(__name__)
PDN_ALL = dict()


class Report:
    def __init__(self, filename: str, purpose_date: datetime.date = None, **options):
        self.order_type = "Основной договор"
        self.filename = str(filename)
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
        # datetime.now().date().replace(day=1) - timedelta(days=1)
        if purpose_date is None:
            self.report_date = datetime.now().date().replace(day=1) - timedelta(days=1)
        else:
            self.report_date = purpose_date
        # self.report_date =  datetime.strptime("30.06.2022","%d.%m.%Y").date() #  datetime.now().date().replace(day=1) - timedelta(days=1)
        self.reference = {}  # ссылки на документы в рамках одного документа
        self.documents = []  # документы по порядку
        self.wa = {}
        self.kategoria = {}
        self.warnings = []
        self.fields = {}
        self.tarifs = [(0, "noname"), (1, "Постоянный"), (2, "Старт")]
        self.discounts = (2, 10, 31, 33, 42, 45, 47, 44, 46, 48, 51, 52)
        self.is_find_columns = False
        self.options = options
        self.errors: List[ErrorReport] = list()

    def get_parser(self):
        if self.read():
            headers = get_columns_head(self.suf)
            for index, self.record in enumerate(self.parser.records):
                if self.is_find_columns is False:
                    self.set_columns(headers)
                    if index > 20:
                        break
                if self.is_find_columns is True:
                    self.__record_order_type()
                    self.__record_client()
                    if self.__get_current_client():
                        self.__record_order_number(index)
                        self.__record_document()
                        self.__record_order_date()
                        self.__record_order_period()
                        self.__record_order_payments()
                        self.__record_order_pdn()
                        self.__record_order_rate()
                        self.__record_order_tarif()
                        self.__record_doc_period()
                        self.__record_doc_summa()
                        self.__set_order_count_days()
        return

    def __record_client(self) -> None:
        num = self.fields.get("FLD_NAME")
        names = re.findall(PATT_NAME, self.record[num])
        if bool(names) is False:
            names = re.findall(PATT_NAME, self.record[num + 1])
        if bool(names) is False:
            if self.fields.get("FLD_DOCUMENT") is not None and self.__is_find(
                PATT_TIME_IN_DOCUMENT, "FLD_DOCUMENT"
            ):
                if self.fields.get("FLD_BEG_DEBET_proc") and bool(
                    self.record[self.fields.get("FLD_BEG_DEBET_proc")].strip()
                ):
                    names = [self.record[num + 1]]
                elif self.fields.get("FLD_BEG_CREDIT_proc") and bool(
                    self.record[self.fields.get("FLD_BEG_CREDIT_proc")].strip()
                ):
                    names = [self.record[num]]
        if names:
            self.current_client_key = re.sub("<...>|\s|\n", "", names[0].lower())
            pattern = f".+(?=: начислено)|.+"
            name = re.sub(f"<...>|{PATT_TIME_IN_DOCUMENT}", "", names[0])
            name = re.search(pattern, name).group(0)
            self.clients.setdefault(
                self.current_client_key,
                Client(name=name),
            )
            self.__set_new_order()
            self.__record_order_summa()

    def __get_current_client(self) -> Client:
        if self.current_client_key:
            return self.clients[self.current_client_key]
        else:
            return None

    def __get_current_order(self, is_cached: bool = False) -> Order:
        client: Client = self.__get_current_client()
        if client:
            return client.order_cache
        else:
            None

    def __get_current_document(self) -> Client:
        client: Client = self.__get_current_client()
        if client:
            if client.documents:
                return client.documents[-1]
        else:
            None

    def __set_current_order(self, order: Order) -> Order:
        for attribute, value in vars(order).items():
            if (
                isinstance(value, str)
                or isinstance(value, float)
                or isinstance(value, int)
                or isinstance(value, Decimal)
                or isinstance(value, datetime)
                or isinstance(value, date)
            ) and value:
                setattr(
                    self.clients[self.current_client_key].order_cache, attribute, value
                )

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
            client.order_cache.client = client
            self.__clean_payment()

    def __set_new_document(self) -> None:
        client: Client = self.__get_current_client()
        if client:
            text = self.record[self.fields.get("FLD_DOCUMENT")]
            debet = (
                self.record[self.fields.get("FLD_BEG_DEBET_proc")].strip()
                if self.fields.get("FLD_BEG_DEBET_proc")
                else ""
            )
            credit = (
                self.record[self.fields.get("FLD_BEG_CREDIT_proc")].strip()
                if self.fields.get("FLD_BEG_CREDIT_proc")
                else ""
            )
            document = Document(text, debet=debet, credit=credit)
            document.order = self.__get_current_order()
            document.client = client
            client.documents.append(document)
            self.documents.append(document)
        return

    def __set_doc_period(self, text: str) -> None:
        document: Document = self.__get_current_document()
        if document:
            period = to_date(text)
            if period:
                document.date_period = period

    def __set_doc_summa(self, text: str) -> None:
        document: Document = self.__get_current_document()
        if document:
            document.summa = float(text)

    def __clean_payment(self) -> None:
        order: Order = self.__get_current_order()
        if order:
            order.payment_cache = Payment()

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
        if self.fields.get("FLD_DOCUMENT") is not None and self.__is_find(
            PATT_TIME_IN_DOCUMENT, "FLD_DOCUMENT"
        ):
            if self.__is_find(PATT_CURRENCY, f"FLD_BEG_DEBET_{self.suf}"):
                date_payment = self.record[self.fields["FLD_DOC_PERIOD"]]
                self.__add_payment(date_payment, "FLD_BEG_DEBET", "B", "D")
            if self.__is_find(PATT_CURRENCY, f"FLD_BEG_CREDIT_{self.suf}"):
                date_payment = self.record[self.fields["FLD_DOC_PERIOD"]]
                self.__add_payment(date_payment, "FLD_BEG_CREDIT", "O", "C")

    def __record_order_date(self):
        order = self.__get_current_order()
        if self.__is_find(PATT_DOG_DATE, "FLD_DATE", "date_order"):
            order.date_order = to_date(
                get_long_date(self.record[self.fields["FLD_DATE"]])
            )
            order.date_begin = order.date_order
            if self.__is_find(PATT_DOG_DATE, "FLD_DATE_BEGIN", "date_begin", True):
                order.date_begin = to_date(
                    get_long_date(self.record[self.fields["FLD_DATE_BEGIN"]])
                )
        if self.__is_find(PATT_DOG_DATE, "FLD_DATE_FROZEN", "date_frozen"):
            order.date_frozen = to_date(
                get_long_date(self.record[self.fields["FLD_DATE_FROZEN"]])
            )
        self.__set_order_field(PATT_DOG_DATE, "FLD_DATE_END", "date_end")
        self.__set_order_field(
            PATT_COUNT_DAYS, "FLD_COUNT_DAYS_DELAY", "count_days_delay"
        )

    def __record_order_pdn(self):
        self.__set_order_field(PATT_PDN, "FLD_PDN", "pdn", True, value_type="float")
        if getattr(self.clients[self.current_client_key].order_cache, "pdn", 0) != 0:
            if self.clients[self.current_client_key].order_cache.pdn < 3:
                self.clients[self.current_client_key].order_cache.pdn *= 100
            PDN_ALL[self.current_client_key] = self.clients[
                self.current_client_key
            ].order_cache.pdn

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
                "FLD_SUMMA_PAYMENT",
                "summa_payment",
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
        self.__set_order_field(
            PATT_CURRENCY,
            f"FLD_END_DEBET_proc_58",
            f"debet_end_proc_58",
            is_forced=False,
            value_type="decimal",
        )
        self.__set_order_field(
            PATT_CURRENCY,
            f"FLD_SUMMA_RESERVE_MAIN",
            f"summa_reserve_main_58",
            is_forced=False,
            value_type="decimal",
        )
        self.__set_order_field(
            PATT_CURRENCY,
            f"FLD_SUMMA_RESERVE_MAIN_PDN",
            f"summa_reserve_main_58_pdn",
            is_forced=False,
            value_type="decimal",
        )
        self.__set_order_field(
            PATT_CURRENCY,
            f"FLD_SUMMA_RESERVE_PROC",
            f"summa_reserve_proc_58",
            is_forced=False,
            value_type="decimal",
        )
        self.__set_order_field(
            PATT_CURRENCY,
            f"FLD_SUMMA_RESERVE_PROC_PDN",
            f"summa_reserve_proc_58_pdn",
            is_forced=False,
            value_type="decimal",
        )
        return

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
        numbers = []
        if self.fields.get("FLD_NUMBER") is not None:
            numbers = re.search(
                PATT_DOG_NUMBER, self.record[self.fields.get("FLD_NUMBER")]
            )
        elif self.fields.get("FLD_DOCUMENT") is not None:
            numbers = re.search(
                PATT_DOG_NUMBER, self.record[self.fields.get("FLD_DOCUMENT")]
            )
        if numbers:
            client: Client = self.__get_current_client()
            if client:
                order_cache = client.order_cache
                self.__set_new_order()
                if order_cache.client is not None and len(client.orders) == 0:
                    self.__set_current_order(order_cache)
                order = self.__get_current_order()
                order.type = self.order_type
                order.name = client.name
                order = self.__get_current_order()
                number = get_order_number(numbers[0])
                order.number = number
                order.row = index
                self.reference.setdefault(order.number, order)
                self.__record_order_summa(True)
                self.__push_current_order()
                return

    # Документ
    def __record_document(self):
        if self.fields.get("FLD_DOCUMENT") is None:
            return
        numbers = re.search(
            PATT_TIME_IN_DOCUMENT, self.record[self.fields.get("FLD_DOCUMENT")]
        )
        if numbers:
            self.__set_new_document()
        return

    # Период
    def __record_doc_period(self):
        if self.fields.get("FLD_DOC_PERIOD") is None:
            return
        numbers = re.search(
            PATT_DOC_PERIOD, self.record[self.fields.get("FLD_DOC_PERIOD")]
        )
        if numbers:
            self.__set_doc_period(self.record[self.fields.get("FLD_DOC_PERIOD")])
        return

    # Сумма документа
    def __record_doc_summa(self):
        if (
            self.fields.get("FLD_END_DEBET_proc") is None
            and self.fields.get("FLD_END_CREDIT_proc") is None
        ):
            return
        numbers = re.search(
            PATT_CURRENCY, self.record[self.fields.get("FLD_END_DEBET_proc")]
        )
        if numbers is None:
            numbers = re.search(
                PATT_CURRENCY, self.record[self.fields.get("FLD_END_CREDIT_proc")]
            )
        if numbers:
            self.__set_doc_summa(numbers.group(0))
        return

    def __set_order_count_days(self, order: Order = None):
        if order is None:
            order = self.__get_current_order()
        if order.count_days == 0:
            order.count_days = 31
        if order and order.count_days != 0:
            try:
                if order.date_begin:
                    is_recalc_proc = False
                    order.date_end = order.date_begin + timedelta(order.count_days)
                    order.count_days_period = self.__get_count_days_in_period(order)
                    order.count_days_common = self.__get_count_days_common(order)
                    order.count_days_delay = self.__get_count_days_delay(order)
                    if order.debet_end_main == 0:
                        is_recalc_proc = True
                        order.debet_end_main = max(order.summa - order.credit_main, 0)
                    order.summa_percent_period = round(
                        order.debet_end_main
                        * Decimal(order.rate)
                        / 100
                        * order.count_days_period,
                        2,
                    )
                    order.summa_percent_all = round(
                        order.debet_end_main
                        * Decimal(order.rate)
                        / 100
                        * order.count_days_common,
                        2,
                    )

                    summa_percent_max = order.summa * Decimal(
                        get_max_margin_rate(order.date_order)
                    )

                    if order.debet_end_proc == 0:
                        if order.debet_end_proc_58 != 0:
                            order.debet_end_proc = order.debet_end_proc_58
                        # else:
                        #     # order.debet_end_proc = round(
                        #     debet_end_proc = round(
                        #         max(
                        #             min(summa_percent_max, order.summa_percent_all)
                        #             - order.summa_payment
                        #             - (
                        #                 order.credit_proc
                        #                 if order.summa_payment == 0
                        #                 else 0
                        #             ),
                        #             0,
                        #         ),
                        #         2,
                        #     )
                        #     if debet_end_proc !=0:
                        #         print("Вычисляем дебет")
                        if (
                            self.options.get("option_is_archi")
                            and is_recalc_proc is False
                        ):
                            if self.__is_find(
                                PATT_CURRENCY,
                                "FLD_SUMMA_RESERVE_PROC_PDN",
                                "summa_reserve_proc_58_pdn",
                                True,
                            ):
                                order.debet_end_proc = order.debet_end_proc_58

                    debet_end_proc = round(
                        max(
                            min(summa_percent_max, order.summa_percent_all)
                            - order.summa_payment
                            - (order.credit_proc if order.summa_payment == 0 else 0),
                            0,
                        ),
                        2,
                    )
                    if debet_end_proc != order.debet_end_proc and order.debet_end_proc == 0:
                        error = ErrorReport()
                        error.number = order.number
                        error.name = order.name
                        error.summa = float(order.debet_end_proc)
                        error.summa_dop_1 = float(debet_end_proc)
                        error.summa_dop_2 = float(order.debet_end_proc - debet_end_proc)
                        error.description = "Остаток по Дт.76 не совпадает с расчетным "
                        self.errors.append(error)

                    order.percent = self.__get_reserve_persent(order)
                    if order.percent == 0:
                        order.calc_reserve_main = round(
                            order.debet_end_main * Decimal(0.1), 2
                        )
                        order.calc_reserve_proc = round(
                            order.debet_end_proc * Decimal(0.1), 2
                        )
                    elif order.percent > 0:
                        order.calc_reserve_main = round(
                            order.debet_end_main * Decimal(order.percent), 2
                        )
                        order.calc_reserve_proc = round(
                            order.debet_end_proc * Decimal(order.percent), 2
                        )

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
        if order.date_frozen:
            last_day_on_period = order.date_frozen
        count_days = (last_day_on_period - order.date_begin).days
        if order.tarif.code in self.discounts:
            if order.tarif.code == 48:
                if count_days >= 16 and count_days <= 22:
                    count_days = 16
                elif count_days > 22:
                    count_days -= 7
            else:
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
            # return (
            #     (
            #         last_day_on_period - order.date_begin - timedelta(order.count_days)
            #     ).days
            #     - ((order.count_days // 31) - 1 if order.count_days >= 31 else 0)
            #     if last_day_on_period > order.date_end
            #     else 0
            # )
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
        if self.fields.get(column_fld_name, -1) != -1 and (
            order_fld_name is None or (not getattr(order, order_fld_name) or is_forced)
        ):
            if is_forced or re.search(
                pattern, self.record[self.fields.get(column_fld_name)], re.IGNORECASE
            ):
                return True
        return False

    def __set_order_field(
        self,
        pattern: str,
        column_fld_name: str = None,
        order_fld_name: str = "",
        is_forced: bool = False,
        value_type: str = "",
    ) -> bool:
        if self.__is_find(pattern, column_fld_name, order_fld_name, is_forced):
            value = (
                self.record[self.fields[column_fld_name]]
                .replace("Нет данных", "")
                .strip()
            )
            order = self.__get_current_order()
            try:
                if value_type == "decimal":
                    if value == "":
                        value = "0"
                    setattr(
                        order,
                        order_fld_name,
                        Decimal(value),
                    )
                elif value_type == "float":
                    if value == "":
                        value = "0"
                    setattr(
                        order,
                        order_fld_name,
                        float(value),
                    )
                else:
                    setattr(order, order_fld_name, value)
                return True
            except Exception as ex:
                pass
        return False

    def __add_payment(
        self, date_payment: str, fld_name: str, p_type: str, p_category: str
    ):
        if self.__is_find(PATT_CURRENCY, f"{fld_name}_{self.suf}"):
            payment: Payment = self.__get_current_payment()
            payment.summa = Decimal(
                self.record[self.fields.get(f"{fld_name}_{self.suf}")]
            )
            payment.date = to_date(get_long_date(date_payment))
            payment.kind = self.suf
            payment.type = p_type
            payment.category = p_category
            if self.fields.get(f"FLD_BEG_DEBET_ACCOUNT_{self.suf}") is not None:
                payment.account_debet = self.record[
                    self.fields.get(f"FLD_BEG_DEBET_ACCOUNT_{self.suf}")
                ]
            if self.fields.get(f"FLD_BEG_CREDIT_ACCOUNT_{self.suf}") is not None:
                payment.account_credit = self.record[
                    self.fields.get(f"FLD_BEG_CREDIT_ACCOUNT_{self.suf}")
                ]

            self.__push_current_payment()

    # Устанавливаем номера колонок
    def set_columns(self, items):
        col = 0
        for cell in self.record:
            if re.search("Оборотно-сальдовая ведомость по счету", cell):
                x = re.findall("(?<=г\. - ).+(?= г[\.])", cell)
                if x:
                    self.report_date = to_date(x[0])

            for item in items:
                for name in item["name"]:
                    if not self.fields.get(name) and re.search(item["pattern"], cell):
                        self.fields[name] = col + item["off_col"]
            col += 1
        if self.fields.get("FLD_NAME", -1) != -1:
            self.is_find_columns = True

    def read(self) -> bool:
        return self.parser.read()

    def write_to_excel(self) -> str:
        file_name = (
            "report_irkom" if self.options.get("option_is_archi") else "report_irkom"
        )
        exel = ExcelExporter(file_name)
        return exel.write(self)

    def check_order_exist(self, x, number):
        order = x.reference.get(number)
        if order:
            order.is_cached = True
            return True
        else:
            return False

    def check_items(self, items):
        for item in items:
            for key, order in item.reference.items():
                if (
                    order.is_cached is False
                    and item.name.find("76") != -1
                    and order.debet_end_proc != 0
                ):
                    error = ErrorReport()
                    error.number = order.number
                    error.name = order.name
                    error.summa_dop_1 = float(order.debet_end_proc)
                    error.description = "Не найден договор в {}".format(item.name)
                    self.errors.append(error)

    def union_all(self, items):
        pattern: re.Pattern = re.compile("_proc|_main|pdn|rate|count_|date|percent")
        if not items:
            return
        order_attrs = [x for x in get_attributes(Order()) if pattern.search(x)]
        client: Client
        order: Order

        for item in items:
            if re.search("63", item.name):
                for key, client in item.clients.items():
                    for order in client.orders:
                        if self.reference.get(order.number) is None:
                            i = 1
                            while self.clients.get(f"{key}_{i}"):
                                i += 1
                            self.clients[f"{key}_{i}"] = client
                            self.reference.setdefault(order.number, order)
        for item in items:
            if re.search("59", item.name):
                for key, client in item.clients.items():
                    if self.clients.get(f"{key}_{1}") and self.clients.get(key) is None:
                        if (
                            len(self.clients[f"{key}_{1}"].orders) != 0
                            and self.clients[f"{key}_{1}"].orders[0].credit_end_main
                            == 0
                        ):
                            self.clients[f"{key}_{1}"].orders[
                                0
                            ].credit_end_main = client.order_cache.credit_end_main

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
                    if self.check_order_exist(x, order.number)
                ]
                order_item: Order
                for order_item in order_items:
                    if order_item.number == order.number:
                        for attr in order_attrs:
                            if not getattr(order, attr) and getattr(order_item, attr):
                                setattr(order, attr, getattr(order_item, attr))
                            if (re.search("^debet[a-z_]+proc$", attr)) and (
                                getattr(order, attr) != getattr(order_item, attr)
                            ):
                                setattr(
                                    order,
                                    attr,
                                    getattr(order, attr) + getattr(order_item, attr),
                                )
                        if order_item.tarif.code != 0:
                            order.tarif = order_item.tarif
                        payment: Payment
                        for payment in order_item.payments_1c:
                            order.payments_1c.append(payment)
                            if (
                                (
                                    order.date_frozen is None
                                    or order.date_frozen < payment.date
                                )
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
                self.__set_order_count_days(order)
        self.check_items(items)

    def fill_from_archi(self, data: dict):
        if not data:
            return
        client: Client = None
        for client in self.clients.values():
            order: Order = Order()
            for order in client.orders:
                if data["order"].get(order.number):
                    order.rate = data["order"][order.number][1]
                    order.count_days = data["order"][order.number][2]
                    order.tarif = Tarif()
                    order.tarif.code = data["order"][order.number][6]
                    order.tarif.name = data["order"][order.number][7]
                    if (
                        data["order"][order.number][9]
                        and data["order"][order.number][10]
                        and bool(client.passport_number) is False
                    ):
                        client.passport_number = "{} {}".format(
                            data["order"][order.number][9],
                            data["order"][order.number][10],
                        )
                if data["payment"].get(order.number):
                    for payment in data["payment"][order.number]:
                        order.payments_base.append(payment)
                self.__set_order_count_days(order)

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
                if period and summa and tarif:
                    calc_period = period - 7 if tarif in self.discounts else period
                    key = f"{tarif_name}_{rate}_{30 if calc_period <= 30 else 31}"
                    data = self.wa.get(key)
                    period = float(period)
                    if not data:
                        # 46 -Друг
                        self.wa[key] = {
                            "parent": [],
                            "stavka": float(rate),
                            "koef": 226.065
                            if tarif in self.discounts
                            else 365 * float(rate) if rate != 0 else 0.8,
                            "period": period,
                            # "period": period  if calc_period <= 30 else 31,
                            "summa_free": 0,
                            "summa": 0,
                            "count": 0,
                            "value": {},
                        }
                    self.wa[key]["parent"].append(
                        {"period": calc_period, "order": order}
                    )
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
        for key_client, client in self.clients.items():
            pdn = PDN_ALL.get(key_client, 0) / 100
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
        pass

    def __get_reserve_persent(self, order):
        count = order.count_days_delay
        if (count <= 7) and (order.pdn < 0.5 or order.debet_end_main < 10000):
            return -1
        elif count <= 7:
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
