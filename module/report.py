import re
import os
import csv
import json
import pathlib
import traceback
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
                pdn = item['pdn'] if item.get('pdn') else 0.3
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
