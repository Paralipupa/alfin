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
from module.reports.errors import write_errors
from alfin.module.reports.payment_handle_reserve import write_payment_reserve
from alfin.module.reports.payment_handle_kassa import write_payment_kassa
from module.reports.reserve import write_reserve
from module.reports.kategoria import write_kategoria
from module.reports.weighted_average import write_result_weighted_average
from module.reports.clients import write_clients

logger = logging.getLogger(__name__)


class ExcelExporter:
    def __init__(self, file_name: str, page_name: str = None):
        self.name = file_name
        self.workbook = None

    def _set_data_xls(self, report):
        dop_name = ""
        if report.options.get("option_weighted_average"):
            dop_name = "_average"
        elif report.options.get("option_reserve"):
            dop_name = "_reserve"
        elif report.options.get("option_kategory"):
            dop_name = "_pdn"
        elif report.options.get("option_handle"):
            dop_name = "_handle"
        elif (
            report.options.get("option_cb_common")
            or report.options.get("option_cb_kassa")
            or report.options.get("option_cb_rs")
        ):
            dop_name = "_CB"
        WritterClass = get_file_write(self.name + dop_name)
        self.workbook = WritterClass(self.name + dop_name)
        if not self.workbook:
            raise Exception(f"file reading error: {self.name+dop_name}")

    def write(self, report) -> str:
        self._set_data_xls(report)
        if report.options.get("option_clients"):
            write_clients(self, report)
        if report.options.get("option_weighted_average"):
            write_result_weighted_average(self, report.wa)
        if report.options.get("option_kategory"):
            write_kategoria(self, report.kategoria)
        if report.options.get("option_reserve"):
            write_reserve(self, report)
        if report.options.get("option_handle"):
            if report.options.get("option_reserve"):
                write_payment_reserve(self, report)
            else:
                write_payment_kassa(self, report)
        if report.options.get("option_cb_common"):
            write_CBank_common(self, report)
        if report.options.get("option_cb_kassa"):
            write_CBank_kassa(self, report)
        if report.options.get("option_cb_rs"):
            write_CBank_rs(self, report)
        if bool(report.errors) is True:
            write_errors(self, report)
        return self.workbook.save()
