from xlwt import Utils, Formula

from module.data import *


# %% Категория
def write_kategoria(self, kategoria):
    self.workbook.addSheet("Категория")

    row = 0
    col = 0
    names = ["1", "2", "3", "4", "5", "6"]
    for name in names:
        self.workbook.write(row, col, name, "align: horiz center")
        col += 1
    row += 1
    col = 0
    nrow_start = len(kategoria.items()) + 3
    pattern_style_5 = "pattern: pattern solid, fore_colour green; font: color yellow;"
    num_format = "#,##0.00"
    pattern_style_3 = "pattern: pattern solid, fore_colour orange; font: color white"
    for key, value in kategoria.items():
        if key != "0":
            self.workbook.write(row, col, key)
            self.workbook.write(row, col + 1, value["title"])
            self.workbook.write(
                row, col + 3, value["count4"], pattern_style_5, num_format
            )
            if value["count4"] > 0:
                self.workbook.write(
                    row,
                    col + 2,
                    Formula(
                        f"SUM({Utils.rowcol_pair_to_cellrange(nrow_start,col+5,nrow_start+value['count4']-1,col+6)})"
                    ),
                    pattern_style_3,
                    num_format,
                )
                self.workbook.write(
                    row,
                    col + 4,
                    Formula(
                        f"SUM({Utils.rowcol_pair_to_cellrange(nrow_start,col+4,nrow_start+value['count4']-1,col+4)})"
                    ),
                    pattern_style_5,
                    num_format,
                )
                s = f"SUMIF({Utils.rowcol_pair_to_cellrange(nrow_start,col+7,nrow_start+value['count4']-1,col+7)};\">90\";{Utils.rowcol_pair_to_cellrange(nrow_start,col+5,nrow_start+value['count4']-1,col+5)})"
                s += f"+SUMIF({Utils.rowcol_pair_to_cellrange(nrow_start,col+7,nrow_start+value['count4']-1,col+7)};\">90\";{Utils.rowcol_pair_to_cellrange(nrow_start,col+6,nrow_start+value['count4']-1,col+6)})"
                self.workbook.write(
                    row, col + 5, Formula(s), pattern_style_3, num_format
                )
                self.workbook.write(
                    row,
                    col + 9,
                    Formula(
                        f"SUM({Utils.rowcol_pair_to_cellrange(nrow_start,col+9,nrow_start+value['count4']-1,col+10)})"
                    ),
                    pattern_style_3,
                    num_format,
                )
            nrow_start += value["count4"] + 1
            row += 1
    self.workbook.write(row, col + 1, "Всего", "align: horiz left")
    if row > 7:
        self.workbook.write(
            row,
            col + 2,
            Formula(f"SUM({Utils.rowcol_pair_to_cellrange(row-7,col+2,row-1,col+2)})"),
            pattern_style_3,
            num_format,
        )
        self.workbook.write(
            row,
            col + 3,
            Formula(f"SUM({Utils.rowcol_pair_to_cellrange(row-7,col+3,row-1,col+3)})"),
            pattern_style_5,
            num_format,
        )
        self.workbook.write(
            row,
            col + 4,
            Formula(f"SUM({Utils.rowcol_pair_to_cellrange(row-7,col+4,row-1,col+4)})"),
            pattern_style_5,
            num_format,
        )
        self.workbook.write(
            row,
            col + 5,
            Formula(f"SUM({Utils.rowcol_pair_to_cellrange(row-7,col+4,row-1,col+5)})"),
            pattern_style_3,
            num_format,
        )

    row += 2
    self.workbook.write(row, col + 3, "ПДН")
    self.workbook.write(row, col + 4, "(5)основная")
    self.workbook.write(row, col + 5, "(3,6)основная")
    self.workbook.write(row, col + 6, "(3,6)процент")
    self.workbook.write(row, col + 7, "Дн.пр.")
    self.workbook.write(row, col + 9, "Кт.59")
    self.workbook.write(row, col + 10, "Кт.63")
    for key, value in kategoria.items():
        row += 1
        self.workbook.write(row, col, key)
        for val in value["items"]:
            self.workbook.write(row, col + 1, val["name"])
            order: Order = Order()
            order = val["parent"]
            self.workbook.write(
                row,
                col + 2,
                Formula(order.link["number_address"])
                if order.link.get("number_address")
                else order.number,
            )
            self.workbook.write(
                row,
                col + 3,
                Formula(order.link["pdn_address"])
                if order.link.get("pdn_address")
                else order.pdn,
            )
            self.workbook.write(
                row,
                col + 4,
                Formula(order.link["summa_address"])
                if order.link.get("summa_address")
                else order.summa,
            )
            self.workbook.write(
                row,
                col + 5,
                Formula(order.link["debet_end_main_address"])
                if order.link.get("debet_end_main_address")
                else order.debet_main,
            )
            self.workbook.write(
                row,
                col + 6,
                Formula(order.link["debet_end_proc_address"])
                if order.link.get("debet_end_proc_address")
                else order.debet_proc,
            )
            self.workbook.write(
                row,
                col + 7,
                Formula(order.link["count_days_delay_address"])
                if order.link.get("count_days_delay_address")
                else order.count_days_delay,
            )
            self.workbook.write(
                row,
                col + 9,
                Formula(order.link["credit_end_main_address"])
                if order.link.get("credit_end_main_address")
                else order.credit_end_main,
            )
            self.workbook.write(
                row,
                col + 10,
                Formula(order.link["credit_end_proc_address"])
                if order.link.get("credit_end_proc_address")
                else order.credit_end_proc,
            )
            row += 1
