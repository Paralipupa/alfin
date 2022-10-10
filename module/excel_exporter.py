from module.file_readers import get_file_write
import csv
import json


class ExcelExporter:

    def __init__(self, file_name: str, page_name: str = None):
        self.name = file_name

    def _get_data_xls(self):
        WritterClass = get_file_write(self.name)
        data_writer = WritterClass(self.name)
        if not data_writer:
            raise Exception(f'file reading error: {self.name}')
        return data_writer

    def write(self, report) -> bool:
        data_excel = self._get_data_xls()
        sh = data_excel.book.add_sheet("Лист 1")
        self.write_docs(sh, report.docs)
        sh = data_excel.book.add_sheet("Лист 2")
        self.write_result_weighted_average(sh, report.result)
        sh = data_excel.book.add_sheet("Лист 3")
        self.write_kategoria(sh, report.kategoria)
        data_excel.save()

    def write_docs(self, sh, docs) -> bool:
        names = [{'name': 'number', 'title': 'Номер', 'type': ''}, {'name': 'date', 'title': '', 'type': 'Дата'},
                 {'name': 'summa', 'title': 'Сумма', 'type': 'float'},
                 {'name': 'proc', 'title': 'Ставка', 'type': 'float'},
                 {'name': 'tarif', 'title': 'Тариф', 'type': ''},
                 {'name': 'period', 'title': 'Срок', 'type': 'int'},
                 {'name': 'beg_debet', 'title': 'Сальдо начальное', 'type': 'float'},
                 {'name': 'turn_debet', 'title': 'Оборот', 'type': 'float'},
                 {'name': 'turn_credit',
                     'title': 'Сальдо конечное дебет', 'type': 'float'},
                 {'name': 'end_debet', 'title': 'Сальдо конечное кредит', 'type': 'float'},
                 {'name': 'pdn', 'title': 'ПДН', 'type': 'float'},
                 {'name': 'summa_deb_common', 'title': 'Долг общий', 'type': 'float'},
                 {'name': 'summa_deb_main', 'title': 'Долг основной', 'type': 'float'}, {
            'name': 'summa_deb_proc', 'title': 'Долг процент', 'type': 'float'}
        ]
        row = 0
        col = 0
        sh.write(row, col, 'ФИО')
        for name in names:
            col += 1
            sh.write(row, col, name['title'])
        row += 1
        for doc in docs:
            col = 0
            for dog in doc['dogovor']:
                sh.write(row, col, doc['name'])
                for name in names:
                    col += 1
                    sh.write(row, col, (float(dog[name['name']]) if name['type'] == 'float' else int(
                        dog[name['name']]) if name['type'] == 'int' else str(dog[name['name']])) if dog.get(name['name']) else None)
                col = 0
                row += 1

    def write_result_weighted_average(self, sh, result):
        names = [{'name': 'stavka'}, {'name': 'period'}, {'name': 'koef'},
                 {'name': 'summa_free'}, {'name': 'summa'}, {'name': 'count'}]
        index = 0
        for key, value in result.items():
            index += 1
            row = 0
            col = (index-1)*5
            if isinstance(value, dict):
                sh.write(row, col, key)
                for name in names:
                    row += 1
                    sh.write(row, col, value[name['name']])
                sorted_value = sorted(
                    value['value'].items(), key=lambda x: float(x[0]))
                row += 1
                for val in sorted_value:
                    row += 1
                    sh.write(row, col, int(val[1]))
                    sh.write(row, col+1, float(val[0]))
                    sh.write(row, col+2, float(val[0])*int(val[1]))
                    sh.write(row, col+3, float(val[0])
                             * (value['koef'])*int(val[1]))
            else:
                row += (index-2)
                sh.write(row, 2, key)
                sh.write(row, 3, value)

    def write_kategoria(self, sh, kategoria):
        row = 0
        col = 0
        for key, value in kategoria.items():
            sh.write(row, col, key)
            sh.write(row, col+2, value['count'])
            sh.write(row, col+3, value['summa'])
            row +=1
