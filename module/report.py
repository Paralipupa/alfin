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

    def read(self):
        self.parser.read()

    def write(self, filename: str = 'output', doc_type: str = 'reference'):
        with open(f'{filename}.json', mode='w', encoding='utf-8') as file:
            if doc_type == 'reference':
                docs = self.reference 
            else:
                docs = self.reference 
            jstr = json.dumps(docs, indent=4,
                            ensure_ascii=False)

            file.write(jstr)

    def set_reference(self):
        for doc in self.docs:
            for item in doc['dogovor']:
                name: str = doc['name'].replace(' ', '').lower()
                number: str = item['number']
                self.reference[f'{name}_{number}'] = item

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
