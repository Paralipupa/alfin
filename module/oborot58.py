import re
import csv
import json
from module.excel_importer import ExcelImporter
from module.report import Report
from module.settings import *

class Oborot58(Report):

    def get_parser(self):
        self.read()
        index = 0
        for rec in self.parser.records:
            index += 1
            if re.search(PATT_NAME, rec[FLD58_NAME], re.IGNORECASE):
                doc = {'name': rec[FLD58_NAME], 'dogovor': []}
                self.docs.append(doc)
            elif re.search(PATT_DOG_NAME, rec[FLD58_NAME], re.IGNORECASE):
                self.docs[-1]['dogovor'].append({})
            elif re.search(PATT_DOG_NUMBER, rec[FLD58_NAME], re.IGNORECASE):
                self.docs[-1]['dogovor'][-1]['number'] = rec[FLD58_NAME]
                self.docs[-1]['dogovor'][-1]['summa'] = rec[FLD58_BEG_DEBET]
                self.docs[-1]['dogovor'][-1]['beg_debet_main'] = rec[FLD58_BEG_DEBET]
                self.docs[-1]['dogovor'][-1]['turn_debet_main'] = rec[FLD58_TURN_DEBET]
                self.docs[-1]['dogovor'][-1]['turn_credit_main'] = rec[FLD58_TURN_CREDIT]
                self.docs[-1]['dogovor'][-1]['end_debet_main'] = rec[FLD58_END_DEBET]

                self.docs[-1]['dogovor'][-1]['row'] = index
                self.checksum['debet'] += float(rec[FLD58_BEG_DEBET]) if rec[FLD58_BEG_DEBET] else 0
                self.checksum['current'] += float(rec[FLD58_TURN_DEBET]) if rec[FLD58_TURN_DEBET] else 0
                self.checksum['credit'] += float(rec[FLD58_END_DEBET]) if rec[FLD58_END_DEBET] else 0
            elif re.search(PATT_DOG_DATE, rec[FLD58_NAME], re.IGNORECASE):
                self.docs[-1]['dogovor'][-1]['date'] = rec[FLD58_NAME]
        self.set_reference()
        self.write('rep_58','docs')
