from module.report import Report
from module.connect import SQLServer
from module.helpers import timing
from module.serializer import serializer, deserializer

calc_cashe = {}

class Calc:
    def __init__(self, files: list, is_archi: bool = False):
        self.main_wa: Report = None
        self.main_res: Report = None
        self.archi_data = None
        self.items: list[Report] = []
        self.is_archi = is_archi
        for file in files:
            if file.find("58RES") != -1:
                self.main_res = Report(files)
            elif file.find("58WA") != -1:
                self.main_wa = Report(file)
            else:
                self.items.append(Report(file))

    def read(self) -> None:
        for rep in self.items:
            rep.get_parser()
        if self.main_wa:
            self.main_wa.get_parser()
            self.main_wa.union_all(self.items)
            self.read_from_archi()
            self.main_wa.fill_from_archi(self.archi_data)
        if self.main_res:
            self.main_res.get_parser()
            self.main_res.union_all(self.items)
            self.read_from_archi()
            self.main_res.fill_from_archi(self.archi_data)

    def read_from_archi(self):
        numbers_file = "numbers.dump"
        data_file = "data.dump"
        numbers = []
        if self.is_archi:
            numbers = self.main_wa.get_numbers()
            numbers_from_dump = deserializer(numbers_file)
            if numbers == numbers_from_dump:
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
        self.report_kategoria()
        self.report_weighted_average()
        return self.write()


