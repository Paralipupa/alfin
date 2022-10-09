import re
import csv
import json
from module.excel_importer import ExcelImporter
from module.report import Report

from module.settings import *


class Irkom(Report):

    def get_parser(self):
        self.read()
        for rec in self.parser.records:
            if (re.search(PATT_DOG_NUMBER, rec[FLDIRK_NUMBER]) and re.search(PATT_NAME, rec[FLDIRK_NAME])
                    and re.search(PATT_DOG_DATE, rec[FLDIRK_DATE])):
                doc = {'name': rec[FLDIRK_NAME], 'dogovor': []}
                self.docs.append(doc)
                self.docs[-1]['dogovor'].append({})
                self.docs[-1]['dogovor'][-1]['number'] = rec[FLDPDN_NUMBER]
                self.docs[-1]['dogovor'][-1]['date'] = rec[FLDPDN_DATE]
                self.docs[-1]['dogovor'][-1]['beg_debet'] = rec[FLDPDN_SUMMA]
                self.docs[-1]['dogovor'][-1]['proc'] = rec[FLDIRK_PROC]
                self.docs[-1]['dogovor'][-1]['tarif'] = rec[FLDIRK_TARIF]
                self.docs[-1]['dogovor'][-1]['passport'] = rec[FLDIRK_PASSPORT]
                self.docs[-1]['dogovor'][-1]['period'] = rec[FLDIRKPERIOD]
                self.docs[-1]['dogovor'][-1]['summa_deb_common'] = rec[FLDIRK_SUMMA_DEB_COMMOT]
                self.docs[-1]['dogovor'][-1]['summa_deb_main'] = rec[FLDIRK_SUMMA_DEB_MAIN]
                self.docs[-1]['dogovor'][-1]['summa_deb_proc'] = rec[FLDIRK_SUMMA_DEB_PROC]
        self.set_reference()
        self.write('rep_irk')
