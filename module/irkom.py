import re
import csv
import json
import datetime
from module.excel_importer import ExcelImporter
from module.report import Report


from module.settings import *


class Irkom(Report):

    def get_parser(self):
        self.read()
        self.set_columns()
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
                # self.docs[-1]['dogovor'][-1]['end_debet_common'] = rec[FLDIRK_SUMMA_DEB_COMMON]
                # self.docs[-1]['dogovor'][-1]['end_debet_main'] = rec[FLDIRK_SUMMA_DEB_MAIN]
                # self.docs[-1]['dogovor'][-1]['end_debet_proc'] = rec[FLDIRK_SUMMA_DEB_PROC]
                # rec[FLDIRK_SUMMA_DEB_FINE]
                # self.docs[-1]['dogovor'][-1]['end_debet_fine'] = 0
                # rec[FLDIRK_SUMMA_DEB_PENAL]
                # self.docs[-1]['dogovor'][-1]['end_debet_penal'] = 0

                finish_date = datetime.datetime.strptime(
                    rec[FLDIRK_DATE], '%d.%m.%Y') + datetime.timedelta(days=float(rec[FLDIRK_PERIOD_COMMON]))
                first_day_of_current_month = datetime.datetime.today().replace(day=1)
                last_day_of_previous_month = first_day_of_current_month - \
                    datetime.timedelta(days=1)

                self.docs[-1]['dogovor'][-1]['date_finish'] = datetime.datetime.strftime(
                    finish_date, '%d.%m.%Y')
                self.docs[-1]['dogovor'][-1]['count_days'] = (
                    last_day_of_previous_month - finish_date).days

                self.checksum['summa'] += float(rec[FLDIRK_SUMMA]
                                                ) if rec[FLDIRK_SUMMA] else 0
                # self.checksum['debet'] += float(
                #     rec[FLDIRK_SUMMA_DEB_COMMON]) if rec[FLDIRK_SUMMA_DEB_COMMON] else 0
                # self.checksum['current'] += float(
                #     rec[FLDIRK_SUMMA_DEB_MAIN]) if rec[FLDIRK_SUMMA_DEB_MAIN] else 0
                # self.checksum['credit'] += float(
                #     rec[FLDIRK_SUMMA_DEB_PROC]) if rec[FLDIRK_SUMMA_DEB_PROC] else 0

        self.set_reference()
        # self.write('rep_irk')

    def set_columns(self):
        rec = self.parser.records[0]
        for col, val in rec.items():
            if re.search('^ФИО', val):
                FLDIRK_NAME = col
            elif re.search('^Первоначальный срок займа$', val):
                FLDIRK_PERIOD = col
            elif re.search('^Общий срок займа$', val):
                FLDIRK_PERIOD_COMMON = col
            elif re.search('^№ договора$', val):
                FLDIRK_NUMBER = col
            elif re.search('^Дата выдачи', val):
                FLDIRK_DATE = col
            elif re.search('^Сумма займа$', val):
                FLDIRK_SUMMA = col
            elif re.search('^Процентная ставка', val):
                FLDIRK_PROC = col
            elif re.search('^Наименование продукта$', val):
                FLDIRK_TARIF = col
            elif re.search('^Общий долг$', val):
                FLDIRK_SUMMA_DEB_COMMON = col
            elif re.search('^Основной долг$', val):
                FLDIRK_SUMMA_DEB_MAIN = col
            elif re.search('^Долг по процентам$', val):
                FLDIRK_SUMMA_DEB_PROC = col
            elif re.search('^Долг по штрафам$', val):
                FLDIRK_SUMMA_DEB_FINE = col
            elif re.search('^Долг по единовременным штрафам$', val):
                FLDIRK_SUMMA_DEB_PENAL = col
