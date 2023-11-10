from module.data import *

def write_CBank_kassa(self, report: dict):
        def __write_head():
            nonlocal row, col
            if report.options.get("option_is_archi"):
                self.workbook.write(row, 0, "ООО 'МКК Баргузин'")
                self.workbook.write(row, 1, "3827059334")
            else:
                self.workbook.write(row, 0, "МКК 'Ирком'")
                self.workbook.write(row, 1, "3808200398")
            row += 1
            self.workbook.write(row, 0, "Дата")
            self.workbook.write(row, 1, "Номер")
            self.workbook.write(row, 2, "ФИО")
            self.workbook.write(row, 3, "Основание")
            self.workbook.write(row, 4, "Cчет")
            self.workbook.write(row, 5, "Приход")
            self.workbook.write(row, 6, "Расход")
            return

        def __write(document: Document):
            nonlocal client, row
            pattern_style_positive = (
                "pattern: pattern solid, fore_colour green; font: color yellow;"
            )
            pattern_style_negative = (
                "pattern: pattern solid, fore_colour red; font: color yellow;"
            )
            num_format = "#,##0.00"
            col = 0
            if document.summa != 0:
                # 0
                self.workbook.write(
                    row, col, document.date_period, num_format_str=r"dd/mm/yyyy"
                )
                col += 1

                # 1
                self.workbook.write(
                    row,
                    col,
                    document.number,
                )
                col += 1
                # 1
                self.workbook.write(
                    row,
                    col,
                    client.name,
                )
                col += 1

                # 2
                self.workbook.write(
                    row,
                    col,
                    document.basis,
                )
                col += 1

                # 3
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
                col += 1
                # 4
                self.workbook.write(
                    row,
                    col,
                    document.summa if document.code == "1" else None,
                    num_format_str=num_format,
                )
                col += 1

                # 5
                self.workbook.write(
                    row,
                    col,
                    document.summa if document.code == "2" else None,
                    num_format_str=num_format,
                )

                row += 1
            return
        
        self.workbook.addSheet("ЦБ(касса)")
        row, col = 0, 0
        __write_head()
        row += 1
        for document in report.documents:
            client = document.client
            __write(document)
