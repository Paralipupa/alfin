from xlwt import Utils, Formula
from module.settings import *

from module.data import *


def write_payment_kassa(self, report: dict):
    def __write_handle():
        nonlocal row
        if report.options.get("option_is_archi"):
            self.workbook.write(row, 0, "МКК 'Ирком'")
            self.workbook.write(row, 1, "3808200398")
        else:
            self.workbook.write(row, 0, "МКК 'Ирком'")
            self.workbook.write(row, 1, "3808200398")
        row += 1
        self.workbook.write(row, 0, "ФИО")
        self.workbook.write(row, 1, "Договор")
        self.workbook.write(row, 2, "в 1С")
        self.workbook.write(row, 3, "Archi")
        self.workbook.write(row, 4, "Сумма начисления")
        self.workbook.write(row, 5, "Дата начисления")
        self.workbook.write(row, 6, "Дебет")
        self.workbook.write(row, 7, "Кредит")
        self.workbook.write(row, 8, "Субконто")
        self.workbook.write(row, 9, "Назначение")
        self.workbook.write(row, 10, "Содержание")
        self.workbook.write(row, 11, "Процент")

    def __write_up():
        def __write(kind: str):
            nonlocal calculation, client, row, col
            self.workbook.write(
                row,
                col,
                client.name,
            )
            self.workbook.write(
                row,
                col + 1,
                order.number,
            )
            self.workbook.write(
                row,
                col + 2,
                calculation[kind]["1c"],
                num_format_str=num_format,
            )
            self.workbook.write(
                row,
                col + 3,
                calculation[kind]["archi"],
                num_format_str=num_format,
            )
            self.workbook.write(
                row,
                col + 4,
                calculation[kind]["archi"]-calculation[kind]["1c"],
                num_format_str=num_format,
                style_string=pattern_style_positive,
            )
            self.workbook.write(
                row,
                col + 5,
                report.report_date,
                num_format_str=r"dd/mm/yyyy",
            )

            self.workbook.write(row, col + 6, "91.02")

            self.workbook.write(
                row,
                col + 7,
                "59" if kind == "main" else "63",
            )
            self.workbook.write(row, col + 8, "'00010")
            self.workbook.write(
                row,
                col + 9,
                "платеж по основному долгу" if kind == "main" else "платеж по процентам",
            )
            self.workbook.write(
                row,
                col + 11,
                order.percent,
            )
            row += 1

        nonlocal client, order, row, col, num_format
        pattern_style_positive = (
            "pattern: pattern solid, fore_colour green; font: color yellow;"
        )
        num_format = "#,##0.00"
        item_1c: Payment = None
        calculation = {"main": {"archi": 0, "1c": 0}, "proc": {"archi": 0, "1c": 0}}
        for item_arch in order.payments_base:
            if item_arch[COL_PAY_ARCHI_ENABLE] == 1:
                if item_arch[COL_PAY_ARCHI_KIND] == 0:
                    calculation["main"]["archi"] += item_arch[COL_PAY_ARCHI_COST]
                elif item_arch[COL_PAY_ARCHI_KIND] == 1:
                    calculation["proc"]["archi"] += item_arch[COL_PAY_ARCHI_COST]

        for item_1c in order.payments_1c:
            if item_1c.category == "C":
                calculation[item_1c.kind]["1c"] += item_1c.summa

        if calculation["main"]["archi"] != calculation["main"]["1c"]:
            __write("main")
        if calculation["proc"]["archi"] != calculation["proc"]["1c"]:
            __write("proc")
        return

    self.workbook.addSheet("ОперацииВручную")
    row, col = 0, 0
    __write_handle()
    row += 1
    num_format = "#,##0.00"
    order: Order = None
    for client in report.clients.values():
        for order in client.orders:
            __write_up()
