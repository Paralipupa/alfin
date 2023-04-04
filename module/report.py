import re
import os
import json
import pathlib
import datetime
import traceback
from itertools import product
from module.excel_importer import ExcelImporter
from module.excel_exporter import ExcelExporter
from module.helpers import to_date, get_date
from module.settings import *

logger = logging.getLogger(__name__)


class Report:
    def __init__(self, filename: str):
        self.name = str(filename)
        self.suf = "proc" if self.name.find("76") != -1 else "main"
        self.parser = ExcelImporter(self.name)
        self.clients = {}  # клиенты
        self.current_client_name = ""
        self.current_dogovor_number = ""
        self.current_dogovor_type = "Основной договор"
        self.report_date = datetime.datetime.now().date().replace(
            day=1
        ) - datetime.timedelta(days=1)
        self.reference = {}  # ссылки на документы в рамках одного документа
        self.wa = {}
        self.kategoria = {}
        self.warnings = []
        self.fields = {}
        self.__clear_dog_data()

    def __clear_dog_data(self):
        self.dogs = {}
        self.current_dogovor_number = ""

    def __record_client(self, record: list):
        if self.fields.get("FLD_NAME", -1) != -1 and re.search(
            PATT_NAME, record[self.fields.get("FLD_NAME")], re.IGNORECASE
        ):
            self.current_client_name = (
                record[self.fields.get("FLD_NAME")].replace(" ", "").upper()
            )
            self.clients.setdefault(
                self.current_client_name,
                {"name": record[self.fields.get("FLD_NAME")], "dogovor": {}},
            )
            self.__clear_dog_data()

    def __record_dog_name(self, record: list):
        if re.search("Договор", record[self.fields.get("FLD_NUMBER")], re.IGNORECASE):
            self.__clear_dog_data()

    def __record_dog_type(self, record: list):
        if self.fields.get("FLD_NUMBER", -1) != -1 and re.search(
            PATT_DOG_TYPE, record[self.fields.get("FLD_NUMBER")], re.IGNORECASE
        ):
            self.current_dogovor_type = record[self.fields.get("FLD_NUMBER")]

    def __record_dog_pay(self, record: list):
        if self.fields.get("FLD_NUMBER", -1) != -1 and re.search(
            PATT_DOG_PLAT, record[self.fields.get("FLD_NUMBER")], re.IGNORECASE
        ):
            self.clients[self.current_client_name]["dogovor"][
                self.current_dogovor_number
            ]["plat"].append({})
            plat = self.clients[self.current_client_name]["dogovor"][
                    self.current_dogovor_number
                ]["plat"][-1]
            plat["date_proc"] = (
                record[self.fields.get("FLD_NUMBER")].replace(PATT_DOG_PLAT, "").strip()
            )
            plat[f"beg_debet_{self.suf}"] = (
                record[self.fields.get(f"FLD_BEG_DEBET_{self.suf}")]
                if self.fields.get(f"FLD_BEG_DEBET_{self.suf}", -1) != -1
                else ""
            )
            plat[f"turn_debet_{self.suf}"] = (
                record[self.fields.get(f"FLD_TURN_DEBET_{self.suf}")]
                if self.fields.get(f"FLD_TURN_DEBET_{self.suf}", -1) != -1
                else ""
            )
            plat[f"turn_credit_{self.suf}"] = (
                record[self.fields.get(f"FLD_TURN_CREDIT_{self.suf}")]
                if self.fields.get(f"FLD_TURN_CREDIT_{self.suf}", -1) != -1
                else ""
            )
            plat[f"end_debet_{self.suf}"] = (
                record[self.fields.get(f"FLD_END_DEBET_{self.suf}")]
                if self.fields.get(f"FLD_END_DEBET_{self.suf}", -1) != -1
                else ""
            )
            # b = plat.get(f"turn_credit_{self.suf}") is not None and plat.get(f"turn_credit_{self.suf}") != ''
            # b = b and plat.get("report_end_calculate") is None
            # if b:
            #     self.clients[self.current_client_name]["dogovor"][
            #         self.current_dogovor_number
            #     ]["report_end_calculate"] = plat.get(
            #         "date_proc"
            #     )

    def __record_dog_date(self, record):
        if (
            self.fields.get("FLD_DATE", -1) != -1
            and not self.dogs.get("date")
            and re.search(
                PATT_DOG_DATE, record[self.fields.get("FLD_DATE")], re.IGNORECASE
            )
        ):
            self.dogs["date"] = record[self.fields.get("FLD_DATE")] + (
                " 0:00:00"
                if record[self.fields.get("FLD_DATE")].find(":") == -1
                else ""
            )
        if (
            self.fields.get("FLD_DATE_FINISH", -1) != -1
            and not self.dogs.get("date_finish")
            and re.search(
                PATT_DOG_DATE, record[self.fields.get("FLD_DATE_FINISH")], re.IGNORECASE
            )
        ):
            self.dogs["date_finish"] = record[self.fields.get("FLD_DATE_FINISH")]
        if (
            self.fields.get("FLD_COUNT_DAYS_DELAY", -1) != -1
            and not self.dogs.get("count_days_delay")
            and re.search(
                PATT_COUNT_DAYS,
                record[self.fields.get("FLD_COUNT_DAYS_DELAY")],
                re.IGNORECASE,
            )
        ):
            self.dogs["count_days_delay"] = record[
                self.fields.get("FLD_COUNT_DAYS_DELAY")
            ]
        if self.current_dogovor_number:
            if self.dogs.get("date"):
                self.clients[self.current_client_name]["dogovor"][
                    self.current_dogovor_number
                ]["date"] = self.dogs["date"]
            if self.dogs.get("date_finish"):
                self.clients[self.current_client_name]["dogovor"][
                    self.current_dogovor_number
                ]["date_finish"] = self.dogs["date_finish"]
            if self.dogs.get("count_days_delay"):
                self.clients[self.current_client_name]["dogovor"][
                    self.current_dogovor_number
                ]["count_days_delay"] = self.dogs["count_days_delay"]

    def __set_dogovor_count_days(self, dogovor: dict):
        if dogovor.get("date") and dogovor.get("period_common"):
            try:
                dogovor["count_days_last_month"] = self.__get_count_days_last_month(
                    dogovor
                )
                dogovor["count_days_common"] = self.__get_count_days_common(dogovor)
                dogovor["count_days_delay"] = self.__get_count_days_delay(dogovor)
            except Exception as ex:
                logger.exception("Finish date:")

    def __get_dogovor_date_begin(self, dogovor) -> datetime.datetime:
        return get_date(dogovor["date"])

    def __get_dogovor_date_end(self, dogovor) -> datetime.datetime:
        date_dog_end = get_date(dogovor["date"])
        if date_dog_end:
            date_dog_end += datetime.timedelta(days=float(dogovor["period_common"]))
            dogovor["date_finish"] = datetime.datetime.strftime(
                date_dog_end, "%d.%m.%Y"
            )
        return date_dog_end

    def __get_dogovor_date_calculate_end(self, dogovor) -> datetime.datetime:
        if dogovor.get("report_end_calculate"):
            return get_date(dogovor["report_end_calculate"])
        else:
            dogovor["report_end_calculate"] = self.report_date
            return self.__get_dogovor_date_end(dogovor)

    def __get_last_period(self) -> tuple:
        first_day_of_current_month = datetime.datetime.today().replace(day=1).date()
        last_day_of_previous_month = first_day_of_current_month - datetime.timedelta(
            days=1
        )
        first_day_of_previous_month = last_day_of_previous_month.replace(day=1)
        return first_day_of_previous_month, last_day_of_previous_month

    def __get_count_days_last_month(self, dogovor):
        n_first = 0
        npoid = dogovor.get("tarif", 0)
        (
            date_first_day_in_month,
            date_last_day_in_month,
        ) = self.__get_last_period()
        date_dog = self.__get_dogovor_date_begin(dogovor)
        # Дата договора в отчетном месяце
        if (date_dog > date_first_day_in_month - datetime.timedelta(days=1)) and (
            date_dog < date_last_day_in_month + datetime.timedelta(days=1)
        ):
            date_first_day_in_month = date_dog
            n_first = 1
        if dogovor.get("report_frost"):
            date_dog_end = get_date(dogovor["report_frost"])
        else:
            date_dog_end = self.__get_dogovor_date_calculate_end(dogovor)
        # Дата договора в отчетном месяце или ранняя заморозка
        if (
            (date_dog_end > date_first_day_in_month - datetime.timedelta(days=1))
            and (date_dog_end < date_last_day_in_month + datetime.timedelta(days=1))
            or (date_dog_end < date_first_day_in_month)
        ):
            date_last_day_in_month = date_dog_end
        num_days = (date_last_day_in_month - date_first_day_in_month).days + 1
        date1 = date2 = date_dog
        if npoid == 10:
            date1 += datetime.timedelta(days=1)
            date2 += datetime.timedelta(days=9)
        elif (
            (npoid == 31)
            or (npoid == 33)
            or (npoid == 42)
            or (npoid == 45)
            or (npoid == 47)
        ):
            date1 += datetime.timedelta(days=1)
            date2 += datetime.timedelta(days=6)
        elif (npoid == 44) or (npoid == 46) or (npoid == 48):
            date1 += datetime.timedelta(days=16)
            date2 += datetime.timedelta(days=6)
        if not (
            (date_first_day_in_month == date1) and (date_first_day_in_month == date2)
        ):
            if (
                date1 < date_first_day_in_month
                and date2 > date_first_day_in_month - datetime.timedelta(days=1)
            ):
                num_days -= (
                    date2 - date_first_day_in_month + datetime.timedelta(days=1)
                ).days
            elif date1 > date_first_day_in_month - datetime.timedelta(
                days=1
            ) and date2 < date_last_day_in_month + datetime.timedelta(days=1):
                num_days -= (date2 - date1 + datetime.timedelta(days=1)).days
            elif (
                date1 > date_first_day_in_month - datetime.timedelta(days=1)
                and date1 < date_last_day_in_month + datetime.timedelta(days=1)
                and date2 > date_last_day_in_month
            ):
                num_days -= (
                    date_last_day_in_month - date1 + datetime.timedelta(days=1)
                ).days
            elif (
                date1 < date_first_day_in_month
                and date1 < date_last_day_in_month > date_last_day_in_month
            ):
                num_days = 0
        num_days -= n_first
        return num_days if num_days > 0 else 0

    def __get_count_days_common(self, dogovor: dict):
        (
            _,
            last_day_of_previous_month,
        ) = self.__get_last_period()
        npoid = dogovor.get("tarif", 0)
        count_days = (
            last_day_of_previous_month - self.__get_dogovor_date_begin(dogovor)
        ).days
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

    def __get_count_days_delay(self, dogovor: dict):
        (
            _,
            last_day_of_previous_month,
        ) = self.__get_last_period()
        return (
            (last_day_of_previous_month - self.__get_dogovor_date_begin(dogovor)).days
            if last_day_of_previous_month > self.__get_dogovor_date_end(dogovor)
            else 0
        )

    def __record_dog_pdn(self, rec):
        if (
            self.fields.get("FLD_PDN", -1) != -1
            and not self.dogs.get("pdn")
            and re.search(PATT_PDN, rec[self.fields.get("FLD_PDN")], re.IGNORECASE)
        ):
            self.dogs["pdn"] = rec[self.fields.get("FLD_PDN")]
        if self.current_dogovor_number:
            if self.dogs.get("pdn"):
                self.clients[self.current_client_name]["dogovor"][
                    self.current_dogovor_number
                ]["pdn"] = self.dogs["pdn"]

    def __record_dog_proc(self, rec):
        if (
            self.fields.get("FLD_PROC", -1) != -1
            and not self.dogs.get("proc")
            and re.search(PATT_PROC, rec[self.fields.get("FLD_PROC")], re.IGNORECASE)
        ):
            self.dogs["proc"] = round(float(rec[self.fields.get("FLD_PROC")]), 2)
        if self.current_dogovor_number:
            if self.dogs.get("proc"):
                self.clients[self.current_client_name]["dogovor"][
                    self.current_dogovor_number
                ]["proc"] = self.dogs["proc"]

    def __record_dog_tarif(self, rec):
        if (
            self.fields.get("FLD_TARIF", -1) != -1
            and not self.dogs.get("tarif")
            and re.search(PATT_TARIF, rec[self.fields.get("FLD_TARIF")], re.IGNORECASE)
        ):
            self.dogs["tarif"] = rec[self.fields.get("FLD_TARIF")]
            self.dogs["tarif_name"] = rec[self.fields.get("FLD_TARIF")]
        if self.current_dogovor_number:
            if self.dogs.get("tarif"):
                self.clients[self.current_client_name]["dogovor"][
                    self.current_dogovor_number
                ]["tarif"] = self.dogs["tarif"]
                self.clients[self.current_client_name]["dogovor"][
                    self.current_dogovor_number
                ]["tarif_name"] = self.dogs["tarif"]

    def __record_dog_summa(self, record: list, is_forced: bool = False):
        if (
            self.fields.get("FLD_SUMMA", -1) != -1
            and (not self.dogs.get("summa") or is_forced)
            and re.search(
                PATT_CURRENCY, record[self.fields.get("FLD_SUMMA")], re.IGNORECASE
            )
        ):
            self.dogs["summa"] = record[self.fields.get("FLD_SUMMA")]
        if (
            self.fields.get(f"FLD_BEG_DEBET_{self.suf}", -1) != -1
            and (not self.dogs.get(f"beg_debet_{self.suf}") or is_forced)
            and re.search(
                PATT_CURRENCY,
                record[self.fields.get(f"FLD_BEG_DEBET_{self.suf}")],
                re.IGNORECASE,
            )
        ):
            self.dogs[f"beg_debet_{self.suf}"] = record[
                self.fields.get(f"FLD_BEG_DEBET_{self.suf}")
            ]
        if (
            self.fields.get(f"FLD_TURN_DEBET_{self.suf}", -1) != -1
            and (not self.dogs.get(f"turn_debet_{self.suf}") or is_forced)
            and re.search(
                PATT_CURRENCY,
                record[self.fields.get(f"FLD_TURN_DEBET_{self.suf}")],
                re.IGNORECASE,
            )
        ):
            self.dogs[f"turn_debet_{self.suf}"] = record[
                self.fields.get(f"FLD_TURN_DEBET_{self.suf}")
            ]
        if (
            self.fields.get(f"FLD_TURN_CREDIT_{self.suf}", -1) != -1
            and (not self.dogs.get(f"turn_credit_{self.suf}") or is_forced)
            and re.search(
                PATT_CURRENCY,
                record[self.fields.get(f"FLD_TURN_CREDIT_{self.suf}")],
                re.IGNORECASE,
            )
        ):
            self.dogs[f"turn_credit_{self.suf}"] = record[
                self.fields.get(f"FLD_TURN_CREDIT_{self.suf}")
            ]
        if (
            self.fields.get(f"FLD_END_DEBET_{self.suf}", -1) != -1
            and (not self.dogs.get(f"end_debet_{self.suf}") or is_forced)
            and re.search(
                PATT_CURRENCY,
                record[self.fields.get(f"FLD_END_DEBET_{self.suf}")],
                re.IGNORECASE,
            )
        ):
            self.dogs[f"end_debet_{self.suf}"] = record[
                self.fields.get(f"FLD_END_DEBET_{self.suf}")
            ]
        if not self.dogs.get("summa") and (
            self.dogs.get("beg_debet_main") or self.dogs.get("turn_credit_main")
        ):
            self.dogs["summa"] = (
                self.dogs.get("beg_debet_main")
                if self.dogs.get("beg_debet_main")
                else self.dogs.get("turn_credit_main")
            )
        if self.current_dogovor_number:
            self.clients[self.current_client_name]["dogovor"][
                self.current_dogovor_number
            ]["summa"] = self.dogs.get("summa", "")
            self.clients[self.current_client_name]["dogovor"][
                self.current_dogovor_number
            ][f"beg_debet_{self.suf}"] = self.dogs.get(f"beg_debet_{self.suf}", "")
            self.clients[self.current_client_name]["dogovor"][
                self.current_dogovor_number
            ][f"turn_debet_{self.suf}"] = self.dogs.get(f"turn_debet_{self.suf}", "")
            self.clients[self.current_client_name]["dogovor"][
                self.current_dogovor_number
            ][f"turn_credit_{self.suf}"] = self.dogs.get(f"turn_credit_{self.suf}", "")
            self.clients[self.current_client_name]["dogovor"][
                self.current_dogovor_number
            ][f"end_debet_{self.suf}"] = self.dogs.get(f"end_debet_{self.suf}", "")

    def __record_dog_period(self, rec):
        if (
            self.fields.get("FLD_PERIOD", -1) != -1
            and not self.dogs.get("period")
            and re.search(
                PATT_PERIOD, rec[self.fields.get("FLD_PERIOD")], re.IGNORECASE
            )
        ):
            self.dogs["period"] = rec[self.fields.get("FLD_PERIOD")]
        if (
            self.fields.get("FLD_PERIOD_COMMON", -1) != -1
            and not self.dogs.get("period_common")
            and re.search(
                PATT_PERIOD, rec[self.fields.get("FLD_PERIOD_COMMON")], re.IGNORECASE
            )
        ):
            self.dogs["period_common"] = rec[self.fields.get("FLD_PERIOD_COMMON")]
        if self.current_dogovor_number:
            if self.dogs.get("period"):
                self.clients[self.current_client_name]["dogovor"][
                    self.current_dogovor_number
                ]["period"] = self.dogs["period"]
            if self.dogs.get("period_common"):
                self.clients[self.current_client_name]["dogovor"][
                    self.current_dogovor_number
                ]["period_common"] = self.dogs["period_common"]

    def __record_dog_number(self, record: list, index: int):
        if re.search(
            PATT_DOG_NUMBER, record[self.fields.get("FLD_NUMBER")], re.IGNORECASE
        ):
            self.current_dogovor_number = (
                f"0{record[self.fields.get('FLD_NUMBER')].strip()}"
                if len(record[self.fields.get("FLD_NUMBER")].strip()) == LEN_DOG_NUMBER
                else record[self.fields.get("FLD_NUMBER")].strip()
            )
            self.clients[self.current_client_name]["dogovor"].setdefault(
                self.current_dogovor_number, {}
            )
            self.clients[self.current_client_name]["dogovor"][
                self.current_dogovor_number
            ].setdefault("plat", [])
            self.clients[self.current_client_name]["dogovor"][
                self.current_dogovor_number
            ]["type"] = self.current_dogovor_type
            self.clients[self.current_client_name]["dogovor"][
                self.current_dogovor_number
            ]["row"] = index
            self.clients[self.current_client_name]["dogovor"][
                self.current_dogovor_number
            ]["number"] = self.current_dogovor_number
            self.reference.setdefault(
                self.current_dogovor_number,
                self.clients[self.current_client_name]["dogovor"][
                    self.current_dogovor_number
                ],
            )
            self.__record_dog_summa(record, True)

    def get_parser(self, data: dict = None):
        self.read()
        self.set_columns()
        if (
            self.fields.get("FLD_NAME", -1) == -1
            or self.fields.get("FLD_NUMBER", -1) == -1
        ):
            return
        index = -1
        for record in self.parser.records:
            index += 1
            self.__record_dog_type(record)
            self.__record_client(record)
            if self.current_client_name:
                self.__record_dog_name(record)
                self.__record_dog_number(record, index)
                self.__record_dog_date(record)
                self.__record_dog_period(record)
                self.__record_dog_pay(record)
                self.__record_dog_pdn(record)
                self.__record_dog_proc(record)
                self.__record_dog_tarif(record)
                self.__set_dogovor_count_days(
                    self.clients[self.current_client_name]["dogovor"]
                )
        # self.set_reference()

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
                "name": ["FLD_PROC"],
                "pattern": "^Общая сумма долга по процентам$",
                "off_col": 1,
            },
            {"name": ["FLD_PROC"], "pattern": "^Процентная ставка", "off_col": 0},
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

    def write(self, filename: str = "output", doc_type: str = "clients"):
        if doc_type == "reference":
            docs = self.reference
        elif doc_type == "wa":
            docs = self.wa
        elif doc_type == "rezerv":
            docs = self.kategoria
        else:
            docs = self.clients
        os.makedirs("output", exist_ok=True)
        with open(
            pathlib.Path("output", f"{filename}.json"),
            mode="w",
            encoding="windows-1251",
        ) as file:
            jstr = json.dumps(docs, indent=4, ensure_ascii=False)
            file.write(jstr)

    def write_to_excel(self, filename: str = "output_full") -> str:
        exel = ExcelExporter("output_excel")
        return exel.write(self)

    def set_reference(self):
        for client in self.clients.values():
            for number, dogovor in client["dogovor"].items():
                self.reference[number] = dogovor

    def union_all(self, items):
        if not items:
            return
        for number, dogovor in self.reference.items():
            for item in items:
                item_dogovor = item.reference.get(number)
                if item_dogovor:
                    for item_dog_attrib in item_dogovor.keys():
                        if not dogovor.get(item_dog_attrib):
                            dogovor[item_dog_attrib] = item_dogovor[item_dog_attrib]
            d1 = to_date(dogovor.get("date", ""))
            d2 = None
            d3 = None
            if dogovor.get("plat"):
                for plat in dogovor["plat"]:
                    if plat.get("date_proc") and (plat.get("turn_credit_proc") or plat.get("turn_credit_main")  ):
                        if d2 is None:
                            d2 = to_date(plat["date_proc"])
                        d3 = to_date(plat["date_proc"])
            if d2 is None:
                d2 = self.report_date
            if d3 is None:
                d3 = self.report_date
            dogovor["report_date"] = d3.strftime("%d.%m.%Y")  # дата последней оплаты
            if not dogovor.get("report_frost"):
                dogovor["report_frost"] = d2.strftime("%d.%m.%Y")  # дата первой оплаты
            else:
                dogovor["report_frost"] = d2.strftime("%d.%m.%Y")
            if (
                not dogovor.get("count_days_delay")
                and not isinstance(d1, str)
                and not isinstance(d2, str)
            ):
                dogovor["count_days_delay"] = (self.report_date - d1).days
        self.write("clients")

    def fill_from_archi(self, data: dict):
        if not data:
            return
        for client in self.clients.values():
            for dogovor in client["dogovor"].values():
                if data["order"].get(dogovor["number"]):
                    dogovor["proc"] = data["order"][dogovor["number"]][1]
                    dogovor["tarif"] = data["order"][dogovor["number"]][6]
                    dogovor["tarif_name"] = data["order"][dogovor["number"]][7]
                    dogovor["period"] = data["order"][dogovor["number"]][2]
                    dogovor["period_common"] = data["order"][dogovor["number"]][2]
                    self.__set_dogovor_count_days(dogovor)
                if data["payment"].get(dogovor["number"]):
                    dogovor["payment"] = data["payment"][dogovor["number"]]

    # средневзвешенная величина
    def set_weighted_average(self):
        for client in self.clients.values():
            for dogovor in client["dogovor"].values():
                period = dogovor.get("period")
                summa = dogovor.get("turn_debet_main")
                tarif = dogovor.get("tarif_name")
                proc = dogovor.get("proc")
                if period and summa and tarif and proc:
                    key = f"{tarif}_{proc}"
                    data = self.wa.get(key)
                    period = float(period)
                    if not data:
                        # 46 -Друг
                        self.wa[key] = {
                            "parent": [],
                            "stavka": float(proc),
                            "koef": 240.194
                            if tarif == "46" or tarif == "48"
                            else 365 * float(proc),
                            "period": period - 7
                            if tarif == "46" or tarif == "48"
                            else period,
                            "summa_free": 0,
                            "summa": 0,
                            "count": 0,
                            "value": {},
                        }
                    self.wa[key]["parent"].append(dogovor)
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
                            f'ср.взвеш: {client["name"]} {dogovor["number"]}  {summa} period:{period} tarif:{tarif} proc:{proc}'
                        )
        summa = 0
        summa_free = 0
        for key, client in self.wa.items():
            summa += client["summa"]
            summa_free += client["summa_free"]
        self.wa["summa_free"] = summa_free
        self.wa["summa"] = summa
        self.wa["summa_wa"] = summa / summa_free if summa_free != 0 else 1

    # категории потребительских займов
    def set_kategoria(self):
        kategoria = {
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
        reserve = {}
        for client in self.clients.values():
            pdn = 0.3
            for dogovor in client["dogovor"].values():
                pdn = float(dogovor["pdn"]) if dogovor.get("pdn") else pdn
            for dogovor in client["dogovor"].values():
                if dogovor.get("turn_debet_main"):
                    dogovor["turn_debet_main"] = (
                        float(dogovor["turn_debet_main"])
                        if dogovor.get("turn_debet_main")
                        else 0
                    )
                    dogovor["turn_debet_proc"] = (
                        float(dogovor["turn_debet_proc"])
                        if dogovor.get("turn_debet_proc")
                        else 0
                    )
                    dogovor["end_debet_main"] = (
                        float(dogovor["end_debet_main"])
                        if dogovor.get("end_debet_main")
                        else 0
                    )
                    dogovor["end_debet_proc"] = (
                        float(dogovor["end_debet_proc"])
                        if dogovor.get("end_debet_proc")
                        else 0
                    )
                    dogovor["end_debet_fine"] = (
                        float(dogovor["end_debet_fine"])
                        if dogovor.get("end_debet_fine")
                        else 0
                    )
                    dogovor["end_debet_penal"] = (
                        float(dogovor["end_debet_penal"])
                        if dogovor.get("end_debet_penal")
                        else 0
                    )
                    dogovor["pdn"] = (
                        float(dogovor["pdn"]) if dogovor.get("pdn") else pdn
                    )
                    dogovor["count_days_delay"] = (
                        int(dogovor["count_days_delay"])
                        if dogovor.get("count_days_delay")
                        else 0
                    )
                    dogovor["rezerv_percent"] = self.__get_rezerv_percent(
                        dogovor["count_days_delay"]
                    )

                    if dogovor["turn_debet_main"] >= 10000:
                        if dogovor["pdn"] <= 0.3:
                            t = "1"
                        elif dogovor["pdn"] <= 0.4:
                            t = "2"
                        elif dogovor["pdn"] <= 0.5:
                            t = "3"
                        elif dogovor["pdn"] <= 0.6:
                            t = "4"
                        elif dogovor["pdn"] <= 0.7:
                            t = "5"
                        elif dogovor["pdn"] <= 0.8:
                            t = "6"
                        else:
                            t = "7"
                    else:
                        t = "0"
                    kategoria[t]["count4"] += 1
                    kategoria[t]["summa5"] += dogovor["turn_debet_main"]
                    kategoria[t]["summa3"] += (
                        dogovor["end_debet_main"] + dogovor["end_debet_proc"]
                    )
                    if (
                        dogovor["count_days_delay"] > 90
                        and (dogovor["end_debet_main"] + dogovor["end_debet_proc"]) > 0
                    ):
                        kategoria[t]["count6"] += 1
                        kategoria[t]["summa6"] += (
                            dogovor["end_debet_main"] + dogovor["end_debet_proc"]
                        )
                    item = {"name": client["name"], "parent": dogovor}
                    kategoria[t]["items"].append(item)

                    reserve.setdefault(
                        str(dogovor["rezerv_percent"]),
                        {
                            "percent": dogovor["rezerv_percent"],
                            "summa_main": 0,
                            "summa_proc": 0,
                            "count": 0,
                            "items": [],
                        },
                    )
                    reserve[str(dogovor["rezerv_percent"])]["summa_main"] += dogovor[
                        "turn_debet_main"
                    ]
                    reserve[str(dogovor["rezerv_percent"])]["summa_proc"] += dogovor[
                        "end_debet_proc"
                    ]
                    reserve[str(dogovor["rezerv_percent"])]["count"] += 1
                    reserve[str(dogovor["rezerv_percent"])]["items"].append(item)

        reserve = sorted(reserve.items())
        for item in reserve:
            item[1]["items"] = sorted(item[1]["items"], key=lambda x: x["name"])

        self.reserve = reserve
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

    def __get_count_days_in_last_period(self):
        pass
