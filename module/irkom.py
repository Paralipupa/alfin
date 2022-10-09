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
                self.docs[-1]['dogovor'][-1]['number'] = rec[FLDIRK_NUMBER]
                self.docs[-1]['dogovor'][-1]['date'] = rec[FLDIRK_DATE]
                self.docs[-1]['dogovor'][-1]['summa'] = rec[FLDIRK_SUMMA]
                self.docs[-1]['dogovor'][-1]['proc'] = rec[FLDIRK_PROC]
                self.docs[-1]['dogovor'][-1]['tarif'] = rec[FLDIRK_TARIF]
                self.docs[-1]['dogovor'][-1]['passport'] = rec[FLDIRK_PASSPORT]
                self.docs[-1]['dogovor'][-1]['period_common'] = rec[FLDIRK_PERIOD_COMMON]
                self.docs[-1]['dogovor'][-1]['period'] = rec[FLDIRK_PERIOD]
                self.docs[-1]['dogovor'][-1]['summa_deb_common'] = rec[FLDIRK_SUMMA_DEB_COMMOT]
                self.docs[-1]['dogovor'][-1]['summa_deb_main'] = rec[FLDIRK_SUMMA_DEB_MAIN]
                self.docs[-1]['dogovor'][-1]['summa_deb_proc'] = rec[FLDIRK_SUMMA_DEB_PROC]
                self.checksum['summa'] += float(rec[FLDIRK_SUMMA]) if rec[FLDIRK_SUMMA] else 0
                self.checksum['debet'] += float(rec[FLDIRK_SUMMA_DEB_COMMOT]) if rec[FLDIRK_SUMMA_DEB_COMMOT] else 0
                self.checksum['current'] += float(rec[FLDIRK_SUMMA_DEB_MAIN]) if rec[FLDIRK_SUMMA_DEB_MAIN] else 0
                self.checksum['credit'] += float(rec[FLDIRK_SUMMA_DEB_PROC]) if rec[FLDIRK_SUMMA_DEB_PROC] else 0

        self.set_reference()
        self.write('rep_irk')
