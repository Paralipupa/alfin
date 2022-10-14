import re
import os
import json
import pathlib
import datetime
import traceback
from module.excel_importer import ExcelImporter
from module.excel_exporter import ExcelExporter
from module.settings import *
class Report:
    def __init__(self, filename: str, type_turn: str = 'main'):
        self.name = filename
        self.suf = type_turn
        self.parser = ExcelImporter(self.name)
        self.docs = []
        self.reference = {}
        self.result = {}
        self.kategoria = {}
        self.checksum = {'summa': 0, 'debet': 0, 'current': 0, 'credit': 0}
        self.warnings = []
        self.__clear_fld()
        self.__clear_dog_data()

    def __clear_fld(self):
        self.FLD_NAME = None
        self.FLD_NUMBER = None
        self.FLD_DATE = None
        self.FLD_SUMMA = None
        self.FLD_BEG_DEBET = None
        self.FLD_TURN_DEBET = None
        self.FLD_TURN_CREDIT = None
        self.FLD_END_DEBET = None
        self.FLD_END_CREDIT = None
        self.FLD_PDN = None
        self.FLD_PROC = None
        self.FLD_TARIF = None
        self.FLD_PERIOD_COMMON = None
        self.FLD_PERIOD = None
        self.FLD_DATE_FINISH = None
        self.FLD_COUNT_DAYS = None
        self.FLD_SUMMA_FINE = None
        self.FLD_SUMMA_PENAL = None

    def __clear_dog_data(self):
        self.dog_date = None
        self.dog_proc = None
        self.dog_tarif = None
        self.dog_period = None
        self.dog_period_common = None
        self.dog_pdn = None
        self.dog_count_days = None
        self.dog_date_finish = None

    def get_parser(self):
        self.read()
        self.set_columns()
        if not self.FLD_NAME or not self.FLD_NUMBER:
            return
        index = -1
        for rec in self.parser.records:
            index += 1
            if re.search(PATT_NAME, rec[self.FLD_NAME], re.IGNORECASE):
                self.docs.append({'name': rec[self.FLD_NAME], 'dogovor': []})
                self.__clear_dog_data()
            if self.FLD_DATE and re.search(PATT_DOG_DATE, rec[self.FLD_DATE], re.IGNORECASE):
                self.dog_date = rec[self.FLD_DATE]
            if self.FLD_PDN and not self.dog_pdn and re.search(PATT_PDN, rec[self.FLD_PDN], re.IGNORECASE):
                self.dog_pdn = rec[self.FLD_PDN]
            if self.FLD_PROC and not self.dog_proc and re.search(PATT_PROC, rec[self.FLD_PROC], re.IGNORECASE):
                self.dog_proc = rec[self.FLD_PROC]
            if self.FLD_TARIF and not self.dog_tarif and re.search(PATT_TARIF, rec[self.FLD_TARIF], re.IGNORECASE):
                self.dog_tarif = rec[self.FLD_TARIF]
            if self.FLD_PERIOD and not self.dog_period and re.search(PATT_PERIOD, rec[self.FLD_PERIOD], re.IGNORECASE):
                self.dog_period = rec[self.FLD_PERIOD]
            if self.FLD_PERIOD_COMMON and not self.dog_period_common and re.search(PATT_PERIOD, rec[self.FLD_PERIOD_COMMON], re.IGNORECASE):
                self.dog_period_common = rec[self.FLD_PERIOD_COMMON]
            if self.FLD_DATE_FINISH and not self.dog_date_finish and re.search(PATT_DOG_DATE, rec[self.FLD_DATE_FINISH], re.IGNORECASE):
                self.dog_date_finish = rec[self.FLD_DATE_FINISH]
            if self.FLD_COUNT_DAYS and not self.dog_count_days and re.search(PATT_COUNT_DAYS, rec[self.FLD_COUNT_DAYS], re.IGNORECASE):
                self.dog_count_days = rec[self.FLD_COUNT_DAYS]
            if re.search(PATT_DOG_NUMBER, rec[self.FLD_NUMBER], re.IGNORECASE):
                self.docs[-1]['dogovor'].append({})
                self.docs[-1]['dogovor'][-1]['row'] = index
                self.docs[-1]['dogovor'][-1]['number'] = rec[self.FLD_NUMBER]
                if self.dog_date: self.docs[-1]['dogovor'][-1]['date'] = self.dog_date
                if self.FLD_SUMMA: self.docs[-1]['dogovor'][-1]['summa'] = rec[self.FLD_SUMMA]
                if self.FLD_BEG_DEBET : self.docs[-1]['dogovor'][-1][f'beg_debet_{self.suf}'] = rec[self.FLD_BEG_DEBET]
                if self.FLD_TURN_DEBET: self.docs[-1]['dogovor'][-1][f'turn_debet_{self.suf}'] = rec[self.FLD_TURN_DEBET]
                if self.FLD_TURN_CREDIT: self.docs[-1]['dogovor'][-1][f'turn_credit_{self.suf}'] = rec[self.FLD_TURN_CREDIT]
                if self.FLD_END_DEBET: self.docs[-1]['dogovor'][-1][f'end_debet_{self.suf}'] = rec[self.FLD_END_DEBET]
                if self.dog_pdn: self.docs[-1]['dogovor'][-1]['pdn'] = self.dog_pdn
                if self.dog_proc: self.docs[-1]['dogovor'][-1]['proc'] = self.dog_proc
                if self.dog_tarif: self.docs[-1]['dogovor'][-1]['tarif'] = self.dog_tarif
                if self.dog_period: self.docs[-1]['dogovor'][-1]['period'] = self.dog_period
                if self.dog_period_common: self.docs[-1]['dogovor'][-1]['period_common'] = self.dog_period_common
                if self.dog_date and self.dog_period_common and not self.dog_count_days:
                    finish_date = datetime.datetime.strptime(
                        self.dog_date, '%d.%m.%Y') + datetime.timedelta(days=float(self.dog_period_common))
                    first_day_of_current_month = datetime.datetime.today().replace(day=1)
                    last_day_of_previous_month = first_day_of_current_month - \
                        datetime.timedelta(days=1)
                    self.docs[-1]['dogovor'][-1]['date_finish'] = datetime.datetime.strftime(
                        finish_date, '%d.%m.%Y')
                    self.docs[-1]['dogovor'][-1]['count_days'] = (
                        last_day_of_previous_month - finish_date).days
        self.set_reference()

    def set_columns(self):
        index = 0
        for rec in self.parser.records:
            for col_name, val in rec.items():
                if re.search('^Счет$', val) or re.search('^ФИО',val):
                    self.FLD_NAME = col_name
                    for col, val in rec.items():
                        if re.search('^Сальдо на начало периода$', val):
                            self.FLD_BEG_DEBET = col
                        elif re.search('^Обороты за период$', val):
                            self.FLD_TURN_DEBET = col
                            self.FLD_SUMMA = col
                            self.FLD_TURN_CREDIT = str(int(col)+1)
                        elif re.search('^Сальдо на конец периода$', val):
                            self.FLD_END_DEBET = col
                        elif re.search('^Первоначальный срок займа$', val):
                            self.FLD_PERIOD = col
                        elif re.search('^Общий срок займа$', val):
                            self.FLD_PERIOD_COMMON = col
                        elif (re.search('^Процентная ставка', val) or re.search('^Ставка$', val) ):
                            self.FLD_PROC = col
                        elif (re.search('^Наименование продукта$', val) or re.search('^Тариф$', val) ):
                            self.FLD_TARIF = col
                        elif (re.search('^Показатель долговой',val) or re.search('^ПДН$',val)  ):
                            self.FLD_PDN = col
                        if re.search('^Счет$', val) or (re.search('^№ заявки$',val) or re.search('^№ договора$', val) ) and not self.FLD_NUMBER:
                            self.FLD_NUMBER = col
                        if (re.search('^Счет$', val) or re.search('^Дата выдачи', val)) and not self.FLD_DATE:
                            self.FLD_DATE = col
                        if (re.search('^Сумма займа$', val) or re.search('^Выданная сумма займа$', val)) and not self.FLD_SUMMA:
                            self.FLD_SUMMA = col
                    return
            index += 1
            if index > 20:
                return

    def read(self):
        self.parser.read()

    def write(self, filename: str = 'output', doc_type: str = 'docs'):
        if doc_type == 'reference':
            docs = self.reference
        elif doc_type == 'result':
            docs = self.result
        elif doc_type == 'kategoria':
            docs = self.kategoria
        else:
            docs = self.docs
        os.makedirs('output', exist_ok=True)
        with open(pathlib.Path('output', f'{filename}.json'), mode='w', encoding='utf-8') as file:
            jstr = json.dumps(docs, indent=4,
                              ensure_ascii=False)
            file.write(jstr)
        with open(pathlib.Path('output', f'{filename}.json'), mode='a', encoding='utf-8') as file:
            jstr = json.dumps(self.checksum, indent=4,
                              ensure_ascii=False)
            file.write(jstr)

    def write_to_excel(self, filename: str = 'output_full') -> str:
        exel = ExcelExporter('output_excel')
        return exel.write(self)

    def set_reference(self):
        for doc in self.docs:
            for item in doc['dogovor']:
                name: str = re.findall(PATT_FAMALY, doc['name'])
                number: str = item['number']
                self.reference[f'{name[0].replace(" ","").lower()}_{number}'] = item

    def union_all(self, *args):
        if not args:
            return
        dog: dict = self.reference
        for key, value in dog.items():
            for item in args:
                idog = item.reference.get(key)
                if idog:
                    if not value.get('found'):
                        value['found'] = True
                    for ikey in idog.keys():
                        if not value.get(ikey):
                            value[ikey] = idog[ikey]
                else:
                    value['found'] = False
                    if value['turn_debet_main']:
                        self.warnings.append(f'не найден {key} в {item.name}')
        self.write('docs')


# средневзвешенная величина

    def set_weighted_average(self):
        for doc in self.docs:
            for item in doc['dogovor']:
                period = item.get('period')
                summa = item.get('turn_debet_main')
                tarif = item.get('tarif')
                proc = item.get('proc')
                if period and summa and tarif and proc:
                    key = f'{tarif}_{proc}'
                    data = self.result.get(key)
                    period = float(period)
                    if not data:
                        self.result[key] = {'parent':item, 'stavka': float(
                            proc), 'koef': 240.194 if tarif == 'Старт' else 365*float(proc), 'period': period-7 if tarif == 'Старт' else period,
                            'summa_free': 0, 'summa': 0, 'count': 0, 'value': {}}
                    s = self.result[key]['value'].get(summa)
                    if not s:
                        self.result[key]['value'][summa] = 1
                    else:
                        self.result[key]['value'][summa] += 1
                    self.result[key]['summa'] += float(summa) * \
                        self.result[key]['koef']
                    self.result[key]['summa_free'] += float(summa)
                    self.result[key]['count'] += 1
                else:
                    if summa:
                        self.warnings.append(f'ср.взвеш: {doc["name"]} {item["number"]}  {summa} period:{period} tarif:{tarif} proc:{proc}')

        summa = 0
        summa_free = 0
        for key, doc in self.result.items():
            summa += doc['summa']
            summa_free += doc['summa_free']
        self.result['summa_free'] = summa_free
        self.result['summa'] = summa
        self.result['summa_wa'] = summa / summa_free if summa_free != 0 else 1

# категории потребительских займов
    def set_kategoria(self):
        data = {'1': {'title':'', 'count4': 0, 'count6': 0, 'summa5': 0, 'summa3': 0, 'summa6': 0, 'items': []},
                '2': {'title':'', 'count4': 0, 'count6': 0, 'summa5': 0, 'summa3': 0, 'summa6': 0, 'items': []},
                '3': {'title':'', 'count4': 0, 'count6': 0, 'summa5': 0, 'summa3': 0, 'summa6': 0, 'items': []},
                '4': {'title':'', 'count4': 0, 'count6': 0, 'summa5': 0, 'summa3': 0, 'summa6': 0, 'items': []},
                '5': {'title':'', 'count4': 0, 'count6': 0, 'summa5': 0, 'summa3': 0, 'summa6': 0, 'items': []},
                '6': {'title':'', 'count4': 0, 'count6': 0, 'summa5': 0, 'summa3': 0, 'summa6': 0, 'items': []},
                '7': {'title':'', 'count4': 0, 'count6': 0, 'summa5': 0, 'summa3': 0, 'summa6': 0, 'items': []}
                }
        for doc in self.docs:
            pdn = 0.3
            for item in doc['dogovor']:
                pdn = float(item['pdn']) if item.get('pdn') else 0.3
            for item in doc['dogovor']:
                if item.get('turn_debet_main'):
                    item['turn_debet_main'] = float(item['turn_debet_main']) if item.get(
                        'turn_debet_main') else 0
                    item['turn_debet_proc'] = float(item['turn_debet_proc']) if item.get(
                        'turn_debet_proc') else 0
                    item['end_debet_main'] = float(item['end_debet_main']) if item.get(
                        'end_debet_main') else 0
                    item['end_debet_proc'] = float(item['end_debet_proc']) if item.get(
                        'end_debet_proc') else 0
                    item['end_debet_fine'] = float(item['end_debet_fine']) if item.get(
                        'end_debet_fine') else 0
                    item['end_debet_penal'] = float(item['end_debet_penal']) if item.get(
                        'end_debet_penal') else 0
                    item['pdn'] = float(item['pdn']) if item.get('pdn') else pdn
                    item['count_days'] = int(item['count_days']) if item.get(
                        'count_days') else 0
                    if item['turn_debet_main'] >= 10000:
                        if item['pdn'] <= 0.3:
                            t = '1'
                        elif item['pdn'] <= 0.4:
                            t = '2'
                        elif item['pdn'] <= 0.5:
                            t = '3'
                        elif item['pdn'] <= 0.6:
                            t = '4'
                        elif item['pdn'] <= 0.7:
                            t = '5'
                        elif item['pdn'] <= 0.8:
                            t = '6'
                        else:
                            t = '7'
                        data[t]['count4'] += 1
                        data[t]['summa5'] += item['turn_debet_main']
                        data[t]['summa3'] += (item['end_debet_main'] + item['end_debet_proc'])
                        if item['count_days'] > 90 and (item['end_debet_main'] + item['end_debet_proc']) > 0:
                            data[t]['count6'] += 1
                            data[t]['summa6'] += (item['end_debet_main'] +
                                                  item['end_debet_proc'])

                        data[t]['items'].append({'name': doc['name'],'parent': item})
        self.kategoria = data
