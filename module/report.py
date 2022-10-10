import re
import csv
import json
import pathlib
from module.excel_importer import ExcelImporter
from module.excel_exporter import ExcelExporter
from module.settings import *


class Report:
    def __init__(self, filename: str):
        self.name = filename
        self.parser = ExcelImporter(self.name)
        self.docs = []
        self.reference = {}
        self.result = {}
        self.kategoria = {}
        self.checksum = {'summa': 0, 'debet': 0, 'current': 0, 'credit': 0}

    def read(self):
        self.parser.read()

    def write(self, filename: str = 'output', doc_type: str = 'reference'):
        if doc_type == 'reference':
            docs = self.reference
        elif doc_type == 'result':
            docs = self.result
        elif doc_type == 'kategoria':
            docs = self.kategoria
        else:
            docs = self.reference
        with open(pathlib.Path('output', f'{filename}.json'), mode='w', encoding='utf-8') as file:
            jstr = json.dumps(docs, indent=4,
                              ensure_ascii=False)
            file.write(jstr)
        with open(pathlib.Path('output', f'{filename}.json'), mode='a', encoding='utf-8') as file:
            jstr = json.dumps(self.checksum, indent=4,
                              ensure_ascii=False)
            file.write(jstr)

    def write_to_excel(self, filename: str = 'output_full'):
        exel = ExcelExporter('output_excel')
        exel.write(self)

    def set_reference(self):
        for doc in self.docs:
            for item in doc['dogovor']:
                name: str = re.findall(PATT_FAMALY, doc['name'])
                # name: str = doc['name'].replace(' ', '').lower()
                number: str = item['number']
                self.reference[f'{name[0].lower()}_{number}'] = item

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
                    value['found'] = True
                    for ikey in idog.keys():
                        if not value.get(ikey):
                            value[ikey] = idog[ikey]
                else:
                    value['found'] = False
                    if value['turn_debet_main']:
                        logger.warning(f'not found {key} в {item.name}')

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

    def set_kategoria(self):
        data = {'1': {'count': 0, 'summa': 0, 'items': []},
                '2': {'count': 0, 'summa': 0, 'items': []},
                '3': {'count': 0, 'summa': 0, 'items': []},
                '4': {'count': 0, 'summa': 0, 'items': []},
                '5': {'count': 0, 'summa': 0, 'items': []},
                '6': {'count': 0, 'summa': 0, 'items': []},
                '7': {'count': 0, 'summa': 0, 'items': []}
                }
        for doc in self.docs:
            for item in doc['dogovor']:
                if item.get('pdn') and item.get('turn_debet_main'):
                    summa_main = float(item['turn_debet_main']) if item.get(
                        'turn_debet_main') else 0
                    summa_proc = float(item['turn_debet_proc']) if item.get(
                        'turn_debet_proc') else 0
                    pdn = float(item['pdn']) if item['pdn'] else 0
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
                        data[t]['count'] += 1
                        data[t]['summa'] += (summa_main+summa_proc)
                        # data[t]['items'].append(f"{doc['name']}_{item['number']}_{summa_main}_{summa_proc}")
                        data[t]['items'].append(
                            {'name': doc['name'], 'number': item['number'], 'main': summa_main, 'proc': summa_proc})
        self.kategoria = data
