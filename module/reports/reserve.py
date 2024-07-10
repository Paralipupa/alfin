import logging
from xlwt import Utils, Formula
from module.data import *

logger = logging.getLogger(__name__)


def write_reserve(self, report: dict):
    def fill_table(nrow_start: int, row: int, col: int):
        nonlocal pattern_style_3, num_format_2
        f = (
            f"COUNTIF({Utils.rowcol_pair_to_cellrange(nrow_start,col,nrow_start+len(report.clients),col)},"
            + f"{Utils.rowcol_to_cell(row,col,col_abs=True)}"
            + ")"
        )
        self.workbook.write(row, col + 2, Formula(f), pattern_style_3, num_format_2)
        f = (
            f"SUMIF({Utils.rowcol_pair_to_cellrange(nrow_start,col,nrow_start+len(report.clients),col)},"
            + f"{Utils.rowcol_to_cell(row,col,col_abs=True)},"
            + f"{Utils.rowcol_pair_to_cellrange(nrow_start,col+3,nrow_start+len(report.clients),col+3)}"
            + ")"
        )
        self.workbook.write(row, col + 3, Formula(f), pattern_style_3, num_format_2)
        f = (
            f"SUMIF({Utils.rowcol_pair_to_cellrange(nrow_start,col,nrow_start+len(report.clients),col)},"
            + f"{Utils.rowcol_to_cell(row,col,col_abs=True)},"
            + f"{Utils.rowcol_pair_to_cellrange(nrow_start,col+4,nrow_start+len(report.clients),col+4)}"
            + ")"
        )
        self.workbook.write(row, col + 4, Formula(f), pattern_style_3, num_format_2)
        f = (
            f"SUMIF({Utils.rowcol_pair_to_cellrange(nrow_start,col,nrow_start+len(report.clients),col)},"
            + f"{Utils.rowcol_to_cell(row,col,col_abs=True)},"
            + f"{Utils.rowcol_pair_to_cellrange(nrow_start,col+5,nrow_start+len(report.clients),col+5)}"
            + ")"
        )
        self.workbook.write(row, col + 5, Formula(f), pattern_style_5, num_format_2)
        f = (
            f"SUMIF({Utils.rowcol_pair_to_cellrange(nrow_start,col,nrow_start+len(report.clients),col)},"
            + f"{Utils.rowcol_to_cell(row,col,col_abs=True)},"
            + f"{Utils.rowcol_pair_to_cellrange(nrow_start,col+6,nrow_start+len(report.clients),col+6)}"
            + ")"
        )
        self.workbook.write(row, col + 6, Formula(f), pattern_style_5, num_format_2)
        return

    self.workbook.addSheet("Резервы")
    row = 0
    col = 0
    names = [
        "Ставка",
        "Кол-во",
        "Дт.58",
        "Дт.76",
        "Дт.59",
        "Дт.63",
    ]
    for name in names:
        self.workbook.write(row, col, name, "align: horiz center")
        col += 2 if name == "Ставка" else 1
    row += 1
    col = 0
    nrow_start = len(report.reserve) + 3
    pattern_style = "font: height 160;"
    pattern_style_5 = "pattern: pattern solid, fore_colour green; font: color yellow;"
    num_format_2 = "#,##0.00"
    num_format_0 = "#,##0"
    pattern_style_3 = "pattern: pattern solid, fore_colour orange; font: color white"
    for value in report.reserve:
        reserve: Reserve = value[1]
        self.workbook.write(row, col, reserve.percent)
        fill_table(nrow_start, row, col)
        row += 1
    # for col in range(4):
    #     f=f"SUM({Utils.rowcol_pair_to_cellrange(row-len(report.reserve)+1,col+3,row-1,col+3)})"
    #     self.workbook.write(row, col+3, Formula(f))
    row += 1
    col = 3
    for name in names[2:]:
        self.workbook.write(row, col, name, "align: horiz center")
        col += 1
    self.workbook.write(row, col, "Дн.пр.")
    row += 1
    col = 0
    nrow_start = 1
    client: Client = Client()
    for client in report.clients.values():
        for order in client.orders:
            try:
                formula_string = order.link.get("reserve_percent_address", "")
                self.workbook.write(
                    row, col, Formula(formula_string)
                )
            except Exception as ex:
                logger.info(f"{ex}: \n {formula_string}")
            self.workbook.write(
                row,
                col + 1,
                client.name,
            )
            self.workbook.write(
                row,
                col + 2,
                order.number,
            )
            self.workbook.write(
                row,
                col + 3,
                (
                    Formula(order.link["debet_end_main_address"])
                    if order.link.get("debet_end_main_address")
                    else order.debet_main
                ),
                num_format_str=num_format_2,
            )
            self.workbook.write(
                row,
                col + 4,
                (
                    Formula(order.link["debet_end_proc_address"])
                    if order.link.get("debet_end_proc_address")
                    else order.debet_proc
                ),
                num_format_str=num_format_2,
            )

            f = order.link.get("calc_reserve_main_address", "")
            try:
                self.workbook.write(row, col + 5, Formula(f), num_format_str=num_format_2)
            except Exception as ex:
                logger.info(f"{ex}: {f}")
            f = order.link.get("calc_reserve_proc_address")
            try:
                self.workbook.write(row, col + 6, Formula(f), num_format_str=num_format_2)
            except Exception as ex:
                logger.info(f"{ex}: {f}")
            m = order.link.get("count_days_delay_address", "")
            f = f'IF({m}=0,"",{m})'
            try:
                self.workbook.write(
                    row,
                    col + 7,
                    (
                        Formula(f)
                        if order.link.get("count_days_delay_address")
                        else order.count_days_delay
                    ),
                    num_format_str=num_format_0,
                )
            except Exception as ex:
                logger.info(f"{ex}: {f}")
            nrow_start += 1
            row += 1
