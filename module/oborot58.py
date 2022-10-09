import re
import csv
import json
from module.excel_importer import ExcelImporter
from module.report import Report
from module.settings import *

class Oborot58(Report):

    def get_parser(self):
        self.read()
        for rec in self.parser.records:
            if (re.search(PATT_CURRENCY, rec[FLD58_BEG_DEBET]) or re.search(PATT_CURRENCY, rec[FLD58_TURN_DEBET])):
                if re.search(PATT_NAME, rec[FLD58_NAME], re.IGNORECASE):
                    doc = {'name': rec[FLD58_NAME], 'dogovor': []}
                    self.docs.append(doc)
                elif re.search(PATT_DOG_NAME, rec[FLD58_NAME], re.IGNORECASE):
                    self.docs[-1]['dogovor'].append({})
                elif re.search(PATT_DOG_NUMBER, rec[FLD58_NAME], re.IGNORECASE):
                    self.docs[-1]['dogovor'][-1]['number'] = rec[FLD58_NAME]
                    self.docs[-1]['dogovor'][-1]['beg_debet'] = rec[FLD58_BEG_DEBET]
                    self.docs[-1]['dogovor'][-1]['turn_debet'] = rec[FLD58_TURN_DEBET]
                    self.docs[-1]['dogovor'][-1]['turn_credit'] = rec[FLD58_TURN_CREDIT]
                    self.docs[-1]['dogovor'][-1]['end_debet'] = rec[FLD58_END_DEBET]
                elif re.search(PATT_DOG_DATE, rec[FLD58_NAME], re.IGNORECASE):
                    self.docs[-1]['dogovor'][-1]['date'] = rec[FLD58_NAME]
        self.set_reference()
        self.write('rep_58')
