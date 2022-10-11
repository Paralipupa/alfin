import re
import os
import csv
import json
import pathlib
from module.excel_importer import ExcelImporter
from module.excel_exporter import ExcelExporter
from module.settings import *


class Report:
    def __init__(self, filename: str, type_file: str = ''):
        self.name = filename
        self.type_file = type_file
        self.parser = ExcelImporter(self.name)
        self.docs = []
        self.reference = {}
        self.result = {}
        self.kategoria = {}
        self.checksum = {'summa': 0, 'debet': 0, 'current': 0, 'credit': 0}
        self.warnings = []

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

    def get_parser(self):
        pass

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
        for item in self.docs:
            for dog in item['dogovor']:
                period = dog.get('period')
                summa = dog.get('turn_debet_main')
                tarif = dog.get('tarif')
                proc = dog.get('proc')
                if period and summa and tarif and proc:
                    key = f'{tarif}_{proc}'
                    data = self.result.get(key)
                    period = int(period)
                    if not data:
                        self.result[key] = {'stavka': float(
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
                        logger.warning(
                            f'ср.взвеш: {item["name"]} {dog["number"]}  {summa} period:{period} tarif:{tarif} proc:{proc}')
        summa = 0
        summa_free = 0
        for key, item in self.result.items():
            summa += item['summa']
            summa_free += item['summa_free']
        self.result['summa_free'] = summa_free
        self.result['summa'] = summa
        self.result['summa_wa'] = summa / summa_free if summa_free != 0 else 1

# категории потребительских займов
    def set_kategoria(self):
        data = {'1': {'count4': 0, 'summa5': 0, 'summa3': 0, 'summa6': 0, 'items': []},
                '2': {'count4': 0, 'summa5': 0, 'summa3': 0, 'summa6': 0, 'items': []},
                '3': {'count4': 0, 'summa5': 0, 'summa3': 0, 'summa6': 0, 'items': []},
                '4': {'count4': 0, 'summa5': 0, 'summa3': 0, 'summa6': 0, 'items': []},
                '5': {'count4': 0, 'summa5': 0, 'summa3': 0, 'summa6': 0, 'items': []},
                '6': {'count4': 0, 'summa5': 0, 'summa3': 0, 'summa6': 0, 'items': []},
                '7': {'count4': 0, 'summa5': 0, 'summa3': 0, 'summa6': 0, 'items': []}
                }
        for doc in self.docs:
            pdn = 0.3
            for item in doc['dogovor']:
                pdn = item['pdn'] if item.get('pdn') else 0.3
            for item in doc['dogovor']:
                if item.get('turn_debet_main'):
                    summa_main = float(item['turn_debet_main']) if item.get(
                        'turn_debet_main') else 0
                    summa_proc = float(item['turn_debet_proc']) if item.get(
                        'turn_debet_proc') else 0
                    summa_end_main = float(item['end_debet_main']) if item.get(
                        'end_debet_main') else 0
                    summa_end_proc = float(item['end_debet_proc']) if item.get(
                        'end_debet_proc') else 0
                    summa_fine = float(item['end_debet_fine']) if item.get(
                        'end_debet_fine') else 0
                    summa_penal = float(item['end_debet_penal']) if item.get(
                        'end_debet_penal') else 0
                    pdn = float(item['pdn']) if item.get('pdn') else pdn
                    count_days = int(item['count_days']) if item.get(
                        'count_days') else 0
                    if summa_main >= 10000:
                        if pdn <= 0.3:
                            t = '1'
                        elif pdn <= 0.4:
                            t = '2'
                        elif pdn <= 0.5:
                            t = '3'
                        elif pdn <= 0.6:
                            t = '4'
                        elif pdn <= 0.7:
                            t = '5'
                        elif pdn <= 0.8:
                            t = '6'
                        else:
                            t = '7'
                        data[t]['count4'] += 1
                        data[t]['summa5'] += (summa_main)
                        data[t]['summa3'] += (summa_end_main + summa_end_proc)
                        if count_days > 90 and (summa_end_main + summa_end_proc) > 0:
                            data[t]['summa6'] += (summa_end_main +
                                                  summa_end_proc)

                        data[t]['items'].append(
                            {'name': doc['name'], 'number': item['number'], 'main': summa_main, 'proc': summa_proc,
                             'end_main': summa_end_main, 'end_proc': summa_end_proc, 'fine': summa_fine, 'penal': summa_penal,
                             'pdn': pdn, 'count_days': count_days})
        self.kategoria = data
