from module.file_readers import get_file_write
from xlwt import Utils, Formula


class ExcelExporter:

    def __init__(self, file_name: str, page_name: str = None):
        self.name = file_name
        self.workbook = None

    def _set_data_xls(self):
        WritterClass = get_file_write(self.name)
        self.workbook = WritterClass(self.name)
        if not self.workbook:
            raise Exception(f'file reading error: {self.name}')

    def write(self, report) -> str:
        self._set_data_xls()
        self.workbook.addSheet("Общий")
        self.write_docs(report.documents)
        self.workbook.addSheet("Ср.взвешенная")
        self.write_result_weighted_average(report.result)
        self.workbook.addSheet("Резервы")
        self.write_kategoria(report.kategoria)
        self.workbook.addSheet("error")
        self.write_errors(report.warnings)
        return self.workbook.save()

    def write_errors(self, errors):
        row = 0
        col = 0
        for item in errors:
            self.workbook.write(row, col, item)
            row += 1

    def write_docs(self, docs) -> bool:
        names = [{'name': 'number', 'title': 'Номер', 'type': ''}, {'name': 'date', 'title': 'Дата', 'type': ''},
                 {'name': 'summa', 'title': 'Сумма', 'type': 'float'},
                 {'name': 'proc', 'title': 'Ставка', 'type': 'float'},
                 {'name': 'tarif', 'title': 'Тариф', 'type': ''},
                 {'name': 'period', 'title': 'Срок', 'type': 'float'},
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
                 {'name': 'period_common', 'title': 'Общий срок', 'type': 'float'},
                 {'name': 'date_finish', 'title': 'Крайняя дата', 'type': ''},
                 {'name': 'count_days', 'title': 'Просрочка', 'type': 'float'},
                 ]
        row = 0
        col = 0
        self.workbook.write(row, col, 'ФИО')
        for name in names:
            col += 1
            self.workbook.write(row, col, name['title'])
        row += 1
        for doc in docs:
            col = 0
            for dog in doc['dogovor']:
                self.workbook.write(row, col, doc['name'])
                for name in names:
                    col += 1
                    self.workbook.write(row, col, (float(dog[name['name']]) if name['type'] == 'float' else int(
                        dog[name['name']]) if name['type'] == 'int' else str(dog[name['name']])) if dog.get(name['name']) else None)
                    dog[name['name'] +
                        '_address'] = f'{self.workbook.sheet.name}!{Utils.rowcol_to_cell(row,col)}'
                col = 0
                row += 1

    def write_result_weighted_average(self, result):
        names = [{'name': 'stavka'}, {'name': 'period'}, {'name': 'koef'},
                 {'name': 'summa_free'}, {'name': 'summa'}, {'name': 'count'}]
        index = 0
        pattern_style = 'pattern: pattern solid, fore_colour green; font: color yellow;'
        pattern_style_sum = 'pattern: pattern solid, fore_colour white; font: color black;'
        num_format = '#,##0.00'
        for key, value in result.items():
            index += 1
            row = 0
            col = (index-1)*5
            if isinstance(value, dict):
                self.workbook.write(row, col, key)
                for name in names:
                    row += 1
                    self.workbook.write(row, col, value[name['name']])
                sorted_value = sorted(
                    value['value'].items(), key=lambda x: float(x[0]))
                row_start = row
                row += 1
                for val in sorted_value:
                    row += 1
                    self.workbook.write(row, col, float(val[1]))
                    self.workbook.write(row, col+1, float(val[0]))
                    self.workbook.write(
                        row, col+2, float(val[0])*float(val[1]))
                    self.workbook.write(row, col+3, float(val[0])
                                        * (value['koef'])*float(val[1]), style_string=pattern_style_sum, num_format_str=num_format)
                self.workbook.write(
                    row+1, col+2, Formula(f"SUM({Utils.rowcol_pair_to_cellrange(row_start+2,col+2,row,col+2)})"), pattern_style, num_format)
                self.workbook.write(
                    row+1, col+3, Formula(f"SUM({Utils.rowcol_pair_to_cellrange(row_start+2,col+3,row,col+3)})"), pattern_style_sum, num_format)
            elif index > 2:
                row += index-2
                self.workbook.write(row, 2, key)
                style_string = pattern_style_sum
                if key == 'summa_free':
                    style_string = 'pattern: pattern solid, fore_colour green; font: color yellow;'
                elif key == 'summa_wa':
                    style_string = 'pattern: pattern solid, fore_colour yellow; font: color red;'
                self.workbook.write(row, 3, value, style_string, num_format)

    def write_kategoria(self, kategoria):
        row = 0
        col = 0
        names = ['1', '2', '3', '4', '5', '6']
        for name in names:
            self.workbook.write(row, col, name, 'align: horiz center')
            # self.workbook.write(row, col, name,"pattern: pattern solid, fore_color yellow; font: color white; align: horiz center")
            col += 1
        row += 1
        col = 0
        nrow_start = len(kategoria.items())+3
        pattern_style_5 = 'pattern: pattern solid, fore_colour green; font: color yellow;'
        num_format = '#,##0.00'
        pattern_style_3 = 'pattern: pattern solid, fore_colour orange; font: color white'
        for key, value in kategoria.items():
            if key != '0':
                self.workbook.write(row, col, key)
                self.workbook.write(row, col+1, value['title'])
                self.workbook.write(
                    row, col+3, value['count4'], pattern_style_5, num_format)
                if value['count4'] > 0:
                    self.workbook.write(
                        row, col+2, Formula(f"SUM({Utils.rowcol_pair_to_cellrange(nrow_start,col+5,nrow_start+value['count4']-1,col+6)})"), pattern_style_3, num_format)
                    self.workbook.write(
                        row, col+4, Formula(f"SUM({Utils.rowcol_pair_to_cellrange(nrow_start,col+4,nrow_start+value['count4']-1,col+4)})"), pattern_style_5, num_format)
                    s = f"SUMIF({Utils.rowcol_pair_to_cellrange(nrow_start,col+7,nrow_start+value['count4']-1,col+7)};\">90\";{Utils.rowcol_pair_to_cellrange(nrow_start,col+5,nrow_start+value['count4']-1,col+5)})"
                    s += f"+SUMIF({Utils.rowcol_pair_to_cellrange(nrow_start,col+7,nrow_start+value['count4']-1,col+7)};\">90\";{Utils.rowcol_pair_to_cellrange(nrow_start,col+6,nrow_start+value['count4']-1,col+6)})"
                    self.workbook.write(row, col+5, Formula(s),
                                        pattern_style_3, num_format)
                nrow_start += value['count4'] + 1
                row += 1
        self.workbook.write(row, col+1, 'Всего', 'align: horiz left')
        if row > 7:
            self.workbook.write(
                row, col+2, Formula(f"SUM({Utils.rowcol_pair_to_cellrange(row-7,col+2,row-1,col+2)})"), pattern_style_3, num_format)
            self.workbook.write(
                row, col+3, Formula(f"SUM({Utils.rowcol_pair_to_cellrange(row-7,col+3,row-1,col+3)})"), pattern_style_5, num_format)
            self.workbook.write(
                row, col+4, Formula(f"SUM({Utils.rowcol_pair_to_cellrange(row-7,col+4,row-1,col+4)})"), pattern_style_5, num_format)
            self.workbook.write(
                row, col+5, Formula(f"SUM({Utils.rowcol_pair_to_cellrange(row-7,col+4,row-1,col+5)})"), pattern_style_3, num_format)

        row += 2
        self.workbook.write(row, col+4, '(5)основная')
        self.workbook.write(row, col+5, '(3,6)основная')
        self.workbook.write(row, col+6, '(3,6)процент')
        self.workbook.write(row, col+7, 'Дней просрочки')
        for key, value in kategoria.items():
            row += 1
            self.workbook.write(row, col, key)
            for val in value['items']:
                self.workbook.write(row, col+1, val['name'])
                self.workbook.write(
                    row, col+2, Formula(val['parent']['number_address']) if val['parent'].get('number_address') else val['parent'].get('number'))
                self.workbook.write(
                    row, col+3, Formula(val['parent']['pdn_address']) if val['parent'].get('pdn_address') else val['parent'].get('pdn'))
                self.workbook.write(
                    row, col+4, Formula(val['parent']['turn_debet_main_address']) if val['parent'].get('turn_debet_main_address') else val['parent'].get('turn_debet_main'))
                self.workbook.write(
                    row, col+5, Formula(val['parent']['end_debet_main_address']) if val['parent'].get('end_debet_main_address') else val['parent'].get('end_debet_main'))
                self.workbook.write(
                    row, col+6, Formula(val['parent']['end_debet_proc_address']) if val['parent'].get('end_debet_proc_address') else val['parent'].get('end_debet_proc'))
                self.workbook.write(
                    row, col+7, Formula(val['parent']['count_days_address']) if val['parent'].get('count_days_address') else val['parent'].get('count_days'))
                if float(val['parent']['end_debet_main']) > 0 or float(val['parent']['end_debet_proc']) > 0:
                    if val['parent']['count_days'] > 0:
                        percent = self.__get_rezerv_percent(
                            int(val['parent']['count_days']))
                        self.workbook.write(row, col+10, percent)
                        self.workbook.write(
                            row, col+11, Formula(f"{Utils.rowcol_to_cell(row,col+5)}*{Utils.rowcol_to_cell(row,col+10)}"))
                        self.workbook.write(
                            row, col+12, Formula(f"{Utils.rowcol_to_cell(row,col+6)}*{Utils.rowcol_to_cell(row,col+10)}"))
                        self.workbook.write(
                            row, col+13, Formula(f"{Utils.rowcol_to_cell(row,col+11)}+{Utils.rowcol_to_cell(row,col+12)}"))
                    elif float(val['parent']['pdn']) > 0.5:
                        self.workbook.write(
                            row, col+8,  Formula(f"{Utils.rowcol_to_cell(row,col+5)}*0.1"))
                        self.workbook.write(
                            row, col+9, Formula(f"{Utils.rowcol_to_cell(row,col+6)}*0.1"))
                row += 1

        if kategoria.get('0'):
            pass

    def __get_rezerv_percent(self, count: int) -> int:
        if count <= 7:
            return 0
        elif count <= 30:
            return 3/100
        elif count <= 60:
            return 10/100
        elif count <= 90:
            return 20/100
        elif count <= 120:
            return 40/100
        elif count <= 180:
            return 50/100
        elif count <= 270:
            return 65/100
        elif count <= 360:
            return 80/100
        else:
            return 99/100

# from xlwt import Utils
# print Utils.rowcol_pair_to_cellrange(2,2,12,2)
# print Utils.rowcol_to_cell(13,2)
#  ws.write(i, 2, xlwt.Formula("$A$%d+$B$%d" % (i+1, i+1)))
