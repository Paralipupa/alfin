import re
import csv
import json
from module.excel_importer import ExcelImporter
from module.settings import *

class Pdn:
    def __init__(self, filename: str):
        self.name = filename
        self.parser = ExcelImporter(self.name)
        self.docs = []
        self.reference = {}

    def read(self):
        self.parser.read()

    def write(self, filename: str = 'output'):
        with open(f'{filename}.json', mode='w', encoding='utf-8') as file:
            jstr = json.dumps(self.docs, indent=4,
                              ensure_ascii=False)
            file.write(jstr)

    def set_reference(self):
        for doc in self.docs:
            for item in doc['dogovor']:
                name: str = doc['name'].replace(' ', '').lower()
                number: str = item['number']
                self.reference[f'{name}_{number}'] = item

    def get_parser(self):
        self.read()
        for rec in self.parser.records:
            if (re.search(PATT_DOG_NUMBER, rec[FLDPDN_NUMBER]) and re.search(PATT_NAME, rec[FLDPDN_NAME]) \
                and re.search(PATT_DOG_DATE, rec[FLDPDN_DATE])):
                doc = {'name': rec[FLDPDN_NAME], 'dogovor': []}
                self.docs.append(doc)
                self.docs[-1]['dogovor'].append({})
                self.docs[-1]['dogovor'][-1]['number'] = rec[FLDPDN_NUMBER]
                self.docs[-1]['dogovor'][-1]['date'] = rec[FLDPDN_DATE]
                self.docs[-1]['dogovor'][-1]['beg_debet'] = rec[FLDPDN_SUMMA]
                self.docs[-1]['dogovor'][-1]['pdn'] = rec[FLDPDN_PDN]
        self.set_reference()
        self.write('rep_pdn')
