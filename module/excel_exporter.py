import datetime
import re
import logging
from xlwt import Utils, Formula, XFStyle
from module.file_readers import get_file_write
from module.helpers import to_date, get_value_attr, get_max_margin_rate
from module.data import *
from module.reports.cb_kassa import write_CBank_kassa
from module.reports.cb_common import write_CBank_common
from module.reports.cb_ras_schet import write_CBank_rs
from module.reports.payment_handle import write_payment
from module.reports.reserve import write_reserve
from module.reports.kategoria import write_kategoria
from module.reports.weighted_average import write_result_weighted_average
from module.reports.clients import write_clients

logger = logging.getLogger(__name__)


class ExcelExporter:
    def __init__(self, file_name: str, page_name: str = None):
        self.name = file_name
        self.workbook = None

    def _set_data_xls(self):
        WritterClass = get_file_write(self.name)
        self.workbook = WritterClass(self.name)
        if not self.workbook:
            raise Exception(f"file reading error: {self.name}")

    def write(self, report) -> str:
        self._set_data_xls()
        write_clients(self, report)
        if report.options.get("option_weighted_average"):
            write_result_weighted_average(self, report.wa)
        if report.options.get("option_reserve"):
            write_kategoria(self, report.kategoria)
            write_reserve(self, report)
        if report.options.get("option_handle"):
            write_payment(self, report)
        if report.options.get("option_cb"):
            write_CBank_common(self, report)
            write_CBank_kassa(self, report)
            write_CBank_rs(self, report)
        return self.workbook.save()
