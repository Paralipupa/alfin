import re
import csv
import json
from module.excel_importer import ExcelImporter
from module.report import Report

from module.settings import *

class Pdn(Report):

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
                self.docs[-1]['dogovor'][-1]['summa'] = rec[FLDPDN_SUMMA]
                self.docs[-1]['dogovor'][-1]['pdn'] = rec[FLDPDN_PDN]
                self.checksum['summa'] += float(rec[FLDPDN_SUMMA]) if rec[FLDPDN_SUMMA] else 0
        self.set_reference()
        self.write('rep_pdn')
