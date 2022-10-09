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
        pass
