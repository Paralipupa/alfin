from module.file_readers import get_file_reader
import csv
import json
import pathlib, os


class ExcelImporter:
    def __init__(self, file_name: str, page_name: str = None):
        self.name = file_name
        self.names = []
        self.page_name = page_name
        self.records = list()
        self.data_reader = None

    def _get_data_xls(self) -> bool:
        ReaderClass = get_file_reader(self.name)
        if ReaderClass is None:
            return False
        self.data_reader = ReaderClass(self.name, self.page_name)
        if not self.data_reader:
            raise Exception(f"file reading error: {self.name}")
        return True

    def _get_names(self, record: list) -> list:
        names = []
        for cell in record:
            if cell:
                nm = dict()
                nm["name"] = str(cell).strip()
                names.append(nm)
        return names

    def _get_record(self, names: list, record: list) -> dict:
        rec = dict()
        index = 0
        for cell in record:
            if index < len(names):
                rec[names[index]["name"]] = str(cell).strip()
            index += 1
        return rec

    def read(self) -> bool:
        if self._get_data_xls():
            self.records = self.data_reader
            return True
        return False

    def write(self, file_output: str = "output") -> bool:
        os.makedirs("output", exist_ok=True)
        with open(
            pathlib.Path("output", f"{file_output}.json"),
            mode="w",
            encoding="windows-1251",
        ) as file:
            jstr = json.dumps(self.records, indent=4, ensure_ascii=False)
            file.write(jstr)

        with open(
            pathlib.Path("output", f"{file_output}.csv"),
            mode="w",
            encoding="windows-1251",
        ) as file:
            names = [x for x in self.records[0]]
            file_writer = csv.DictWriter(
                file, delimiter=";", lineterminator="\r", fieldnames=names
            )
            file_writer.writeheader()
            for rec in self.records:
                file_writer.writerow(rec)

        return True


if __name__ == "__main__":
    parser = ExcelImporter("input.xls")
    parser.read()
    parser.write("output")
