from module.data import *
from module.error_report import ErrorReport


def write_errors(self, report: dict):
    def __write_head():
        nonlocal row, col
        if report.options.get("option_is_archi"):
            self.workbook.write(row, 0, "МКК 'Ирком'")
            self.workbook.write(row, 1, "3808200398")
        else:
            self.workbook.write(row, 0, "МКК 'Ирком'")
            self.workbook.write(row, 1, "3808200398")
        row += 1
        self.workbook.write(row, 0, "ФИО")
        self.workbook.write(row, 1, "Номер")
        self.workbook.write(row, 2, "Сумма в отчете")
        self.workbook.write(row, 3, "Описание")
        self.workbook.write(row, 4, "Сумма расчетная")
        self.workbook.write(row, 5, "")
        return

    def __write(err: ErrorReport):
        nonlocal client, row
        num_format = "#,##0.00"
        pattern_style = (
            "pattern: pattern solid, fore_colour yellow; font: color black; "
        )
        col = 0
        self.workbook.write(row, col, err.name)
        col += 1
        self.workbook.write(row, col, err.number)
        col += 1
        self.workbook.write(row, col, err.summa, num_format_str=num_format)
        col += 1
        if err.description:
            self.workbook.write(row, col, err.description)
        col += 1
        if err.summa_dop_1:
            if err.summa == 0:
                self.workbook.write(
                    row,
                    col,
                    err.summa_dop_1,
                    num_format_str=num_format,
                    style_string=pattern_style,
                )
            else:
                self.workbook.write(
                    row, col, err.summa_dop_1, num_format_str=num_format
                )
        col += 1
        if err.summa_dop_2:
            self.workbook.write(row, col, err.summa_dop_2, num_format_str=num_format)
        row += 1
        return

    self.workbook.addSheet("errors")
    row, col = 0, 0
    __write_head()
    row += 1
    client: Client = None
    for err in report.errors:
        __write(err)
