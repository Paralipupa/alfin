from xlwt import Utils, Formula

from module.data import *

def write_payment_reserve(self, report: dict):
    
    def __write_handle():
        nonlocal row
        if report.options.get("option_is_archi"):
            self.workbook.write(row, 0, "ООО 'МКК Баргузин'")
            self.workbook.write(row, 1, "3827059334")
        else:
            self.workbook.write(row, 0, "МКК 'Ирком'")
            self.workbook.write(row, 1, "3808200398")
        row += 1
        self.workbook.write(row, 0, "ФИО")
        self.workbook.write(row, 1, "Договор")
        self.workbook.write(row, 2, "в 1С")
        self.workbook.write(row, 3, "Расчет")
        self.workbook.write(row, 4, "Сумма начисления")
        self.workbook.write(row, 5, "Дата начисления")
        self.workbook.write(row, 6, "Дебет")
        self.workbook.write(row, 7, "Кредит")
        self.workbook.write(row, 8, "Субконто")
        self.workbook.write(row, 9, "Назначение")
        self.workbook.write(row, 10, "Содержание")
        self.workbook.write(row, 11, "Процент")
    
    def __write_down(sub: str):
        nonlocal client, order, row, col, num_format
        pattern_style_negative = (
            "pattern: pattern solid, fore_colour red; font: color yellow;"
        )
        num_format = "#,##0.00"
        summa_1c = getattr(order, f"credit_end_{sub}")
        if summa_1c != 0:
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
                summa_1c,
                num_format_str=num_format,
            )
            self.workbook.write(
                row,
                col + 3,
                0,
                num_format_str=num_format,
            )

            self.workbook.write(
                row,
                col + 4,
                summa_1c,
                num_format_str=num_format,
                style_string=pattern_style_negative,
            )
            self.workbook.write(
                row,
                col + 5,
                report.report_date,
                num_format_str=r"dd/mm/yyyy",
            )
            self.workbook.write(row, col + 6, "59" if sub == "main" else "63")
            self.workbook.write(
                row,
                col + 7,
                "91.01",
            )
            self.workbook.write(row, col + 8, "'00010")
            self.workbook.write(
                row,
                col + 9,
                "корректировка резерва",
            )
            self.workbook.write(
                row,
                col + 11,
                order.percent,
            )
            row += 1
        return

    def __write_up(sub: str):
        nonlocal client, order, row, col, num_format
        pattern_style_positive = (
            "pattern: pattern solid, fore_colour green; font: color yellow;"
        )
        num_format = "#,##0.00"
        summa_calc = getattr(order, f"calc_reserve_{sub}")
        if summa_calc != 0:
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
                0,
                num_format_str=num_format,
            )
            self.workbook.write(
                row,
                col + 3,
                summa_calc,
                num_format_str=num_format,
            )
            self.workbook.write(
                row,
                col + 4,
                summa_calc,
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

            f = f"IF("
            f += f'IF({Utils.rowcol_to_cell(row,col+2,col_abs=True)}="",0,{Utils.rowcol_to_cell(row,col+2,col_abs=True)})>'
            f += f'IF({Utils.rowcol_to_cell(row,col+3,col_abs=True)}="",0,{Utils.rowcol_to_cell(row,col+3,col_abs=True)}),'
            f += f'"91.01", ' + ('"59"' if sub == "main" else '"63"')
            f += ")"
            self.workbook.write(
                row,
                col + 7,
                "59" if sub == "main" else "63",
            )
            self.workbook.write(row, col + 8, "'00010")
            self.workbook.write(
                row,
                col + 9,
                "резерв по основному долгу"
                if sub == "main"
                else "резерв по процентам",
            )
            self.workbook.write(
                row,
                col + 11,
                order.percent,
            )
            row += 1
        return

    def __write(sub: str):
        nonlocal client, order, row, col, num_format
        pattern_style_positive = (
            "pattern: pattern solid, fore_colour green; font: color yellow;"
        )
        pattern_style_negative = (
            "pattern: pattern solid, fore_colour red; font: color yellow;"
        )
        num_format = "#,##0.00"
        summa_1c = getattr(order, f"credit_end_{sub}")
        summa_calc = getattr(order, f"calc_reserve_{sub}")
        if summa_1c - summa_calc != 0:
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
                summa_1c,
                num_format_str=num_format,
            )
            self.workbook.write(
                row,
                col + 3,
                summa_calc,
                num_format_str=num_format,
            )
            f = f"ABS({summa_1c-summa_calc})"
            self.workbook.write(
                row,
                col + 4,
                Formula(f),
                num_format_str=num_format,
                style_string=pattern_style_positive
                if summa_1c <= summa_calc
                else pattern_style_negative,
            )
            self.workbook.write(
                row,
                col + 5,
                report.report_date,
                num_format_str=r"dd/mm/yyyy",
            )

            f = f"IF("
            f += f'IF({Utils.rowcol_to_cell(row,col+2,col_abs=True)}="",0,{Utils.rowcol_to_cell(row,col+2,col_abs=True)})<='
            f += f'IF({Utils.rowcol_to_cell(row,col+3,col_abs=True)}="",0,{Utils.rowcol_to_cell(row,col+3,col_abs=True)}),'
            f += f'"91.02", ' + ('"59"' if sub == "main" else '"63"')
            f += ")"
            self.workbook.write(row, col + 6, Formula(f))

            f = f"IF("
            f += f'IF({Utils.rowcol_to_cell(row,col+2,col_abs=True)}="",0,{Utils.rowcol_to_cell(row,col+2,col_abs=True)})>'
            f += f'IF({Utils.rowcol_to_cell(row,col+3,col_abs=True)}="",0,{Utils.rowcol_to_cell(row,col+3,col_abs=True)}),'
            f += f'"91.01", ' + ('"59"' if sub == "main" else '"63"')
            f += ")"
            self.workbook.write(
                row,
                col + 7,
                Formula(f),
            )
            self.workbook.write(row, col + 8, "'00010")

            f = f"IF("
            f += f'IF({Utils.rowcol_to_cell(row,col+2,col_abs=True)}="",0,{Utils.rowcol_to_cell(row,col+2,col_abs=True)})>'
            f += f'IF({Utils.rowcol_to_cell(row,col+3,col_abs=True)}="",0,{Utils.rowcol_to_cell(row,col+3,col_abs=True)}),'
            f += f'"корректировка резерва", ' + (
                '"резерв по основному долгу"'
                if sub == "main"
                else '"резерв по процентам"'
            )
            f += ")"
            self.workbook.write(
                row,
                col + 9,
                Formula(f),
            )
            self.workbook.write(
                row,
                col + 11,
                order.percent,
            )
            row += 1
        return
    
    self.workbook.addSheet("ОперацииВручную")
    row, col = 0, 0
    __write_handle()
    row += 1
    num_format = "#,##0.00"
    for client in report.clients.values():
        for order in client.orders:
            __write_down("main")
    for client in report.clients.values():
        for order in client.orders:
            __write_up("main")
    for client in report.clients.values():
        for order in client.orders:
            __write("proc")

