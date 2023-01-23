from module.file_readers import get_file_write
from xlwt import Utils, Formula, XFStyle
import datetime


class ExcelExporter:

    @staticmethod
    def to_date(x: str):
        patts = ['%d-%m-%Y', '%d.%m.%Y', '%d/%m/%Y', '%Y-%m-%d',
                 '%d-%m-%y', '%d.%m.%y', '%d/%m/%y', '%B %Y']
        d = None
        for p in patts:
            try:
                d = datetime.datetime.strptime(x.split(' ')[0], p)
                return d.date()
            except:
                pass
        return x

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
        self.write_clients(report.clients)
        self.workbook.addSheet("Ср.взвешенная")
        self.write_result_weighted_average(report.wa)
        self.workbook.addSheet("Резервы")
        self.write_kategoria(report.rezerv)
        self.workbook.addSheet("error")
        self.write_errors(report.warnings)
        return self.workbook.save()

    def write_errors(self, errors):
        row = 0
        col = 0
        for item in errors:
            self.workbook.write(row, col, item)
            row += 1

    def write_clients(self, clients) -> bool:
        names = [{'name': 'number', 'title': 'Номер', 'type': '', 'col': 0},
                 {'name': 'date', 'title': 'Дата', 'type': 'date', 'col': 1},
                 {'name': 'summa', 'title': 'Сумма', 'type': 'float', 'col': 2},
                 {'name': 'proc', 'title': 'Ставка', 'type': 'float', 'col': 3},
                 {'name': 'tarif_name', 'title': 'Тариф', 'type': '', 'col': 4},
                 {'name': 'period', 'title': 'Срок', 'type': 'float', 'col': 5},
                 #  {'name': 'beg_debet_main',
                 #      'title': 'Начальная сумма', 'type': 'float', 'col': 6},
                 {'name': 'turn_debet_main', 'title': 'Сальдо нач.',
                     'type': 'float', 'col': 7},
                 {'name': 'turn_credit_main', 'title': 'Кредит',
                     'type': 'float', 'col': 8},
                 {'name': 'end_debet_main', 'title': 'Сальдо кон.',
                     'type': 'float', 'col': 9},
                 #  {'name': 'turn_debet_proc', 'title': 'Процент дебет', 'type': 'float', 'col': 10},
                 #  {'name': 'turn_credit_proc',
                 #      'title': 'Процент кредит', 'type': 'float', 'col': 11},
                 #  {'name': 'end_debet_proc',
                 #      'title': 'Процент остаток', 'type': 'float', 'col': 12},
                 {'name': 'pdn', 'title': 'ПДН', 'type': 'float', 'col': 13},
                 #  {'name': 'period_common', 'title': 'Общий срок',
                 #      'type': 'float', 'col': 14},
                 #  {'name': 'date_finish', 'title': 'Крайняя дата',
                 #      'type': 'date', 'col': 15},
                 #  {'name': 'beg_debet_proc', 'title': 'Долг', 'type': 'float', 'col': 17},
                 {'name': 'turn_debet_proc', 'title': 'Начисление',
                     'type': 'float', 'col': 18},
                 {'name': 'turn_credit_proc', 'title': 'Оплата',
                     'type': 'float', 'col': 19},
                 {'name': 'end_debet_proc', 'title': 'Остаток платежа',
                     'type': 'float', 'col': 20},
                 {'name': 'date_proc', 'title': 'Дата платежа',
                     'type': 'date', 'col': 21},
                 {'name': 'count_days', 'title': 'Просрочка',
                     'type': 'int', 'col': 16},
                 ]
        plat = [{'name': 'date', 'title': 'Дата', 'type': 'string', 'col': 1},
                {'name': 'beg_debet', 'title': 'Остаток', 'type': 'float', 'col': 6},
                {'name': 'turn_debet', 'title': 'Процент дебет',
                    'type': 'float', 'col': 7},
                {'name': 'turn_credit',
                 'title': 'Процент кредит', 'type': 'float', 'col': 8},
                {'name': 'end_debet',
                 'title': 'Процент остаток', 'type': 'float', 'col': 9},
                ]
        row = 0
        col = 0
        self.workbook.write(row, col, 'ФИО')
        for name in names:
            col += 1
            self.workbook.write(row, col, name['title'])
        row += 1
        for client in clients.values():
            for dog in client['dogovor'].values():
                col = 0
                self.workbook.write(row, col, client['name'])
                dog['name' +
                    '_address'] = f'{self.workbook.sheet.name}!{Utils.rowcol_to_cell(row,col)}'
                for name in names:
                    col += 1
                    try:
                        if dog.get(name['name']):
                            value = float(dog[name['name']]) if name['type'] == 'float' else (
                                int(dog[name['name']]) if name['type'] == 'int' else (
                                    self.to_date(dog[name['name']]) if name['type'] == 'date' else (
                                        str(dog[name['name']]) if dog.get(name['name']) else None)))
                            self.workbook.write(
                                row, col, value, num_format_str=r'dd/mm/yyyy' if name['type'] == 'date' else None)
                    except Exception as ex:
                        print(
                            f"{self.workbook.sheet.name} ({name['name']}): {row}, {name['col']}, {value}")
                    dog[name['name'] +
                        '_address'] = f'{self.workbook.sheet.name}!{Utils.rowcol_to_cell(row,col)}'
                if dog.get('plat'):
                    self.workbook.write(row, col, self.to_date(
                        dog['plat'][-1]['date_proc']), num_format_str=r'dd/mm/yyyy')
                    dog['plat'][-1]['date_proc']
                row += 1

    def write_result_weighted_average(self, result):
        if len(result) == 0:
            return
        names = [{'name': 'stavka', 'title':'Ставка'}, {'name': 'period', 'title':'Срок'}, {'name': 'koef', 'title':'Коэфф.'},]
        pattern_style = 'pattern: pattern solid, fore_colour green; font: color yellow;'
        pattern_style_wa = 'pattern: pattern solid, fore_colour yellow; font: color black;'
        num_format = '#,##0.00'
        index = 0
        for key, value in result.items():
            index += 1
            row = 0
            col = (index-1)*5+1
            if isinstance(value, dict):
                self.workbook.write(row, col-1, key.split("_")[0])
                for name in names:
                    row += 1
                    self.workbook.write(row, col-1, name['title'])
                    self.workbook.write(row, col, value[name['name']])                    
                self.workbook.write(row+1, col-1, 'Сумма')
                self.workbook.write(row+1, col, Formula(f"SUM({Utils.rowcol_pair_to_cellrange(row+5,col+2,row+5+len(value['value'])-1,col+2)})"), num_format_str=num_format)
                self.workbook.write(row+2, col-1, 'Сумма(ср.вз.)')
                self.workbook.write(row+2, col, Formula(f"SUM({Utils.rowcol_pair_to_cellrange(row+5,col+3,row+5+len(value['value'])-1,col+3)})"), num_format_str=num_format)
                self.workbook.write(row+3, col-1, 'Кол-во')
                self.workbook.write(row+3, col, Formula(f"SUM({Utils.rowcol_pair_to_cellrange(row+5,col,row+5+len(value['value'])-1,col)})"))
                sorted_value = sorted(
                    value['value'].items(), key=lambda x: float(x[0]))
                row += 3
                row_start = row
                row += 1
                for val in sorted_value:
                    row += 1
                    self.workbook.write(row, col, float(val[1]))
                    self.workbook.write(row, col+1, float(val[0]))
                    self.workbook.write(
                        row, col+2, Formula(f"{Utils.rowcol_to_cell(row,col)}*{Utils.rowcol_to_cell(row,col+1)}"))
                    self.workbook.write(
                        row, col+3, Formula(f"{Utils.rowcol_to_cell(row,col+2)}*{Utils.rowcol_to_cell(3,col)}"))
                self.workbook.write(
                    row+1, col+2, Formula(f"SUM({Utils.rowcol_pair_to_cellrange(row_start+2,col+2,row,col+2)})"), num_format_str=num_format)
                self.workbook.write(
                    row+1, col+3, Formula(f"SUM({Utils.rowcol_pair_to_cellrange(row_start+2,col+3,row,col+3)})"), num_format_str=num_format)
                row += 2
                for dog in value['parent']:
                    self.workbook.write(row, col, Formula(dog['name_address']) if dog.get(
                        'name_address') else dog.get('name',''))
                    self.workbook.write(row, col+1, Formula(dog['number_address']) if dog.get(
                        'number_address') else dog.get('number',''))
                    self.workbook.write(row, col+2, Formula(dog['summa_address']) if dog.get(
                        'summa_address') else dog.get('summa',''), num_format_str=num_format)
                    row += 1
        self.workbook.write(0, 2, 'Общая сумма')
        self.workbook.write(0, 3, Formula(f"SUM({Utils.rowcol_pair_to_cellrange(4,0,4,len(result)*5)})"), num_format_str=num_format)
        self.workbook.write(1, 2, 'Общая сумма(ср.вз.)')
        self.workbook.write(1, 3, Formula(f"SUM({Utils.rowcol_pair_to_cellrange(5,0,5,len(result)*5)})"), num_format_str=num_format)
        self.workbook.write(2, 2, 'Сред.взвеш.')
        self.workbook.write(2, 3, Formula(f"SUM({Utils.rowcol_pair_to_cellrange(3,0,3,len(result)*5)})/COUNT({Utils.rowcol_pair_to_cellrange(3,0,3,len(result)*5)})"), style_string=pattern_style_wa)

    def write_kategoria(self, kategoria):
        row = 0
        col = 0
        names = ['1', '2', '3', '4', '5', '6']
        for name in names:
            self.workbook.write(row, col, name, 'align: horiz center')
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
        self.workbook.write(
            row, 13, Formula(f"SUM({Utils.rowcol_pair_to_cellrange(11,13,nrow_start+value['count4']-1,13)})"), pattern_style_5, num_format)
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
        self.workbook.write(row-1, col+13, 'Резервы')
        self.workbook.write(row, col+10, 'Процент (просрочки)')
        self.workbook.write(row, col+11, 'Сумма (основной)')
        self.workbook.write(row, col+12, 'Сумма (процент)')
        self.workbook.write(row, col+13, 'Итого')
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
                            row, col+11, Formula(f"{Utils.rowcol_to_cell(row,col+5,row_abs=True)}*{Utils.rowcol_to_cell(row,col+10)}") )                            
                        self.workbook.write(
                            row, col+12, Formula(f"{Utils.rowcol_to_cell(row,col+6)}*{Utils.rowcol_to_cell(row,col+10)}"))
                        self.workbook.write(
                            row, col+13, Formula(f"{Utils.rowcol_to_cell(row,col+11)}+{Utils.rowcol_to_cell(row,col+12)}"))
                    elif float(val['parent']['pdn']) > 0.5 and key != '0':
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
