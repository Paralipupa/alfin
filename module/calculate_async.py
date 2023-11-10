import logging
from multiprocessing import Pool, Manager
from datetime import datetime
from module.report import Report
from module.connect import SQLServer
from module.helpers import timing
from module.serializer import serializer, deserializer

calc_cashe = {}
logger = logging.getLogger(__name__)


class Calc:
    def __init__(self, files: list, purpose_date: datetime.date, **options):
        self.main_wa: Report = None
        self.main_res: Report = None
        self.archi_data = None
        self.items: list[Report] = []
        self.options = options
        for file in files:
            if file.find("58рез") != -1 and self.main_res is None:
                self.main_res = Report(file, purpose_date)
                self.main_res.options = self.options
            elif file.find("58рез") != -1 and self.main_wa is None:
                self.main_wa = Report(file, purpose_date)
                self.main_wa.options = self.options
            elif file.find("58") != -1 and self.main_res is None:
                self.main_res = Report(file, purpose_date)
                self.main_res.options = self.options
            elif file.find("58") != -1 and self.main_wa is None:
                self.main_wa = Report(file, purpose_date)
                self.main_wa.options = self.options
            else:
                self.items.append(Report(file, purpose_date=purpose_date, **options))

    def read(self) -> None:
        pool = Pool()
        for rep in self.items:
            pool.apply_async(rep.get_parser())
        if self.main_wa:
            pool.apply_async(self.main_wa.get_parser())
        if self.main_res:
            pool.apply_async(self.main_res.get_parser())
        pool.close()
        pool.join()

        if self.main_res is None and self.main_wa is None and self.items:
            self.main_res = self.items.pop(0)

        if self.main_wa:
            self.main_wa.union_all(self.items)
            self.read_from_archi()
            self.main_wa.fill_from_archi(self.archi_data)
        if self.main_res:
            self.main_res.union_all(self.items)
            self.read_from_archi()
            self.main_res.fill_from_archi(self.archi_data)

    def read_from_archi(self):
        logger.info("Чтение данных из MSSQL")

        numbers_file = "numbers.dump"
        data_file = "data.dump"
        numbers = []
        if self.options.get("option_is_archi"):
            if self.main_wa is not None:
                numbers = self.main_wa.get_numbers()
            elif self.main_res is not None:
                numbers = self.main_res.get_numbers()
            numbers_from_dump = deserializer(numbers_file)
            if len(numbers) != 0 and numbers == numbers_from_dump:
                data = deserializer(data_file)
                self.archi_data = data
            else:
                q = SQLServer()
                if q.connection:
                    self.archi_data = q.get_data_from_archi(numbers)
                    serializer(numbers, numbers_file)
                    serializer(self.archi_data, data_file)

    def report_weighted_average(self) -> None:
        if self.main_wa:
            self.main_wa.set_weighted_average()
            if self.main_res:
                self.main_res.wa = self.main_wa.wa
        elif self.main_res:
            self.main_res.set_weighted_average()

    def report_kategoria(self) -> None:
        if self.main_res:
            self.main_res.set_kategoria()
            if self.main_wa:
                self.main_wa.kategoria = self.main_res.kategoria
        elif self.main_wa:
            self.main_wa.set_kategoria()

    def write(self) -> str:
        return (
            self.main_res.write_to_excel()
            if self.main_res
            else (self.main_wa.write_to_excel() if self.main_wa else None)
        )

    def get_numbers(self):
        if self.main_wa:
            pass
        elif self.main_res:
            pass

    @timing(start_message="Начало")
    def run(self) -> str:
        self.read()
        logger.info("Формирование отчетов")
        self.report_kategoria()
        self.report_weighted_average()
        logger.info("Запись в MS Excel")
        return self.write()
