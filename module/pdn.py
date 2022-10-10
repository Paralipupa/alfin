import re
import csv
import json
from module.excel_importer import ExcelImporter
from module.report import Report

from module.settings import *

class Pdn(Report):

    def get_parser(self):
        self.read()
        self.set_columns()
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

    def set_columns(self):
        rec = self.parser.records[0]
        for col, val in rec.items():
            if re.search('^ФИО',val):
                FLDPDN_NAME = col
            elif  re.search('^Показатель долговой',val):
                FLDPDN_PDN = col
            elif  re.search('^№ заявки$',val):
                FLDPDN_NUMBER = col
            elif  re.search('^Дата подачи',val):
                FLDPDN_DATE = col
            elif  re.search('^Выданная сумма',val):
                FLDPDN_SUMMA = col
