from module.data import *


def write_CBank_rs(self, report: dict):
    def __write_head():
        nonlocal row
        if report.options.get("option_is_archi"):
            self.workbook.write(row, 0, "ООО 'МКК Баргузин'")
            self.workbook.write(row, 1, "3827059334")
        else:
            self.workbook.write(row, 0, "МКК 'Ирком'")
            self.workbook.write(row, 1, "3808200398")
        row += 1
        col = 0
        self.workbook.write(row, col, "Дата")
        col += 1
        self.workbook.write(row, col, "Приход")
        col += 1
        self.workbook.write(row, col, "Расход")
        col += 1
        self.workbook.write(row, col, "Валюта")
        col += 1
        self.workbook.write(row, col, "Наименование контрагента")
        col += 1
        self.workbook.write(row, col, "Номер счета")
        col += 1
        self.workbook.write(row, col, "ИНН")
        col += 1
        self.workbook.write(row, col, "БИК")
        col += 1
        self.workbook.write(row, col, "Содержание операции")
        col += 1
        self.workbook.write(row, col, "Счет корр")
        col += 1
        return

    def __write(document: Document):
        nonlocal client, row
        num_format = "#,##0.00"
        col = 0
        if document.summa != 0:
            # 0
            self.workbook.write(
                row, col, document.date_period, num_format_str=r"dd/mm/yyyy"
            )
            col += 1

            # 8
            self.workbook.write(
                row,
                col,
                document.summa if document.code == "1" else None,
                num_format_str=num_format,
            )
            col += 1

            # 9
            self.workbook.write(
                row,
                col,
                document.summa if document.code == "2" else None,
                num_format_str=num_format,
            )
            col += 1

            # 2
            self.workbook.write(
                row,
                col,
                "РУБ",
            )
            col += 1

            # 3
            self.workbook.write(
                row,
                col,
                client.name,
            )
            col += 1

            # 4
            self.workbook.write(
                row,
                col,
                "",
            )
            col += 1

            # 5
            self.workbook.write(
                row,
                col,
                "",
            )
            col += 1

            # 6
            self.workbook.write(
                row,
                col,
                "",
            )
            col += 1

            # 7
            self.workbook.write(
                row,
                col,
                document.basis,
            )
            col += 1


            # 10
            self.workbook.write(
                row,
                col,
                document.order.payments_1c[0].get_account(
                    document.order.payments_1c[0].account_credit
                )
                if document.code == "1"
                and document.order
                and document.order.payments_1c
                else (
                    document.order.payments_1c[0].get_account(
                        document.order.payments_1c[0].account_debet
                    )
                    if document.code == "2"
                    and document.order
                    and document.order.payments_1c
                    else ""
                ),
            )

            row += 1
        return

    self.workbook.addSheet("ЦБ(рс)")
    row, col = 0, 0
    __write_head()
    row += 1
    for document in report.documents:
        client = document.client
        __write(document)
