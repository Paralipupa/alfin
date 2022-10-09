import re
import csv
import json
from module.excel_importer import ExcelImporter
from module.settings import *


class Report:
    def __init__(self, filename: str):
        self.name = filename
        self.parser = ExcelImporter(self.name)
        self.docs = []
        self.reference = {}
        self.result = {}
        self.checksum = {'summa': 0, 'debet': 0, 'current': 0, 'credit': 0}

    def read(self):
        self.parser.read()

    def write(self, filename: str = 'output', doc_type: str = 'reference'):
        if doc_type == 'reference':
            docs = self.reference
        elif doc_type == 'result':
            docs = self.result
        else:
            docs = self.reference
        with open(f'{filename}.json', mode='w', encoding='utf-8') as file:
            jstr = json.dumps(docs, indent=4,
                              ensure_ascii=False)
            file.write(jstr)
        with open(f'{filename}.json', mode='a', encoding='utf-8') as file:
            jstr = json.dumps(self.checksum, indent=4,
                              ensure_ascii=False)
            file.write(jstr)

    def write_full_csv(self, filename: str = 'output_full'):
        data = []
        for item in self.docs:
            for dog in item['dogovor']:
                data.append(dog.copy())
                data[-1]['name'] = item['name']
        # with open(f'{filename}.csv', mode='w', encoding=ENCONING) as file:
        #     names = [x for x in self.docs[0]['dogovor'][0].keys()]
        #     names.append('name')
        #     file_writer = csv.DictWriter(file, delimiter=";",
        #                                     lineterminator="\r", fieldnames=names)
        #     file_writer.writeheader()
        #     for rec in data:
        #         file_writer.writerow(rec)

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
                    if value['turn_debet']:
                        logger.warning(f'not found {key} в {item.name}')

# средневзвешенная величина
    def weighted_average(self):
        for item in self.docs:
            for dog in item['dogovor']:
                period = dog.get('period')
                summa = dog.get('turn_debet')
                tarif = dog.get('tarif')
                proc = dog.get('proc')
                if period and summa and tarif and proc:
                    key = f'{tarif}_{proc}'
                    data = self.result.get(key)
                    if not data:
                        self.result[key] = {'stavka': float(
                            proc), 'koef': 240.194 if tarif == 'Старт' else 365*float(proc), 'summa_free': 0, 'summa': 0, 'count': 0, 'value': {}}
                    period = int(period)
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
                        logger.warning(f'ср.взвеш: {item["name"]} {dog["number"]}  {summa} period:{period} tarif:{tarif} proc:{proc}')
        summa = 0
        summa_free = 0
        for key, item in self.result.items():
            summa += item['summa']
            summa_free += item['summa_free']
        self.result['summa'] = summa
        self.result['summa_free'] = summa_free
        self.result['summa_wa'] = summa / summa_free if summa_free != 0 else 1
