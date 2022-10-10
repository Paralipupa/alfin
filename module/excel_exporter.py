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

    def write(self, report) -> str:
        data_excel = self._get_data_xls()
        sh = data_excel.book.add_sheet("Лист 1")
        self.write_docs(sh, report.docs)
        sh = data_excel.book.add_sheet("Лист 2")
        self.write_result_weighted_average(sh, report.result)
        sh = data_excel.book.add_sheet("Лист 3")
        self.write_kategoria(sh, report.kategoria)
        return data_excel.save()


    def write_docs(self, sh, docs) -> bool:
        names = [{'name': 'number', 'title': 'Номер', 'type': ''}, {'name': 'date', 'title': 'Дата', 'type': ''},
                 {'name': 'summa', 'title': 'Сумма', 'type': 'float'},
                 {'name': 'proc', 'title': 'Ставка', 'type': 'float'},
                 {'name': 'tarif', 'title': 'Тариф', 'type': ''},
                 {'name': 'period', 'title': 'Срок', 'type': 'int'},
                 {'name': 'beg_debet_main',
                     'title': 'Начальная сумма', 'type': 'float'},
                 {'name': 'turn_debet_main', 'title': 'Дебет', 'type': 'float'},
                 {'name': 'turn_credit_main', 'title': 'Кредит', 'type': 'float'},
                 {'name': 'end_debet_main', 'title': 'Остаток', 'type': 'float'},
                 {'name': 'turn_debet_proc', 'title': 'Процент дебет', 'type': 'float'},
                 {'name': 'turn_credit_proc',
                     'title': 'Процент кредит', 'type': 'float'},
                 {'name': 'end_debet_proc',
                     'title': 'Процент остаток', 'type': 'float'},
                 {'name': 'pdn', 'title': 'ПДН', 'type': 'float'},
                 {'name': 'period_common', 'title': 'Общий срок', 'type': 'int'},
                 {'name': 'date_finish', 'title': 'Крайняя дата', 'type': ''},
                 {'name': 'count_days', 'title': 'Просрочка', 'type': 'int'},
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
        names = ['1', '2', '3', '4', '5', '6']
        for name in names:
            sh.write(row, col, name)
            col += 1
        row += 1
        col = 0
        for key, value in kategoria.items():
            sh.write(row, col, key)
            sh.write(row, col+2, value['summa3'])
            sh.write(row, col+3, value['count4'])
            sh.write(row, col+4, value['summa5'])
            sh.write(row, col+5, value['summa6'])
            row += 1

        row += 1
        sh.write(row, col+3, '(5)')
        sh.write(row, col+5, '(3)')
        for key, value in kategoria.items():
            row += 1
            sh.write(row, col, key)
            for val in value['items']:
                sh.write(row, col+1, val['name'])
                sh.write(row, col+2, val['number'])
                sh.write(row, col+3, float(val['main']))
                sh.write(row, col+4, float(val['pdn']))
                sh.write(row, col+5, float(val['end_main']))
                sh.write(row, col+6, float(val['end_proc']))
                if float(val['end_main']) > 0 and float(val['end_proc']) > 0:
                    if float(val['pdn']) > 0.5:
                        sh.write(row, col+8, float(val['end_main'])*0.1)
                        sh.write(row, col+9, float(val['end_proc'])*0.1)
                    if val['count_days'] > 0:
                        summa_main, summa_proc, percent = self.__summa_rezerv(
                            int(val['count_days']), float(val['end_main']), float(val['end_proc']))
                        sh.write(row, col+11, int(val['count_days']))
                        sh.write(row, col+12, percent)
                        sh.write(row, col+13, summa_main)
                        sh.write(row, col+14, summa_proc)
                        sh.write(row, col+15, summa_main+summa_proc)
                row += 1

    def __summa_rezerv(self, count: int, summa_main: float, summa_proc: float) -> tuple:
        if count <= 7:
            percent = 0
        elif count <= 30:
            percent = 3
        elif count <= 60:
            percent = 10
        elif count <= 90:
            percent = 20
        elif count <= 120:
            percent = 40
        elif count <= 180:
            percent = 50
        elif count <= 270:
            percent = 65
        elif count <= 360:
            percent = 80
        else:
            percent = 99
        summa_main = round((summa_main) * percent/100, 2)
        summa_proc = round((summa_proc) * percent/100, 2)
        return summa_main, summa_proc, percent
