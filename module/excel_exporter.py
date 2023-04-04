import datetime
from xlwt import Utils, Formula, XFStyle
from module.file_readers import get_file_write
from module.helpers import to_date


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
        self.workbook.addSheet("Общий")
        self.write_clients(report)
        self.workbook.addSheet("Ср.взвешенная")
        self.write_result_weighted_average(report.wa)
        self.workbook.addSheet("Категория")
        self.write_kategoria(report.kategoria)
        self.workbook.addSheet("Резервы")
        self.write_reserve(report)
        self.workbook.addSheet("Платежи")
        self.write_payment(report)
        self.workbook.addSheet("error")
        self.write_errors(report.warnings)
        return self.workbook.save()

    def write_errors(self, errors):
        row = 0
        col = 0
        for item in errors:
            self.workbook.write(row, col, item)
            row += 1

    def write_clients(self, report) -> bool:
        names = [
            {"name": "number", "title": "Номер", "type": "", "col": 0},
            {"name": "date", "title": "Дата", "type": "date", "col": 1},
            {"name": "summa", "title": "Сумма", "type": "float", "col": 2},
            {"name": "proc", "title": "Ставка", "type": "float", "col": 3},
            {"name": "tarifs", "title": "Тариф", "type": "", "col": 4},
            {"name": "period", "title": "Срок", "type": "float", "col": 5},
            #  {'name': 'beg_debet_main',
            #      'title': 'Начальная сумма', 'type': 'float', 'col': 6},
            # {
            #     "name": "turn_debet_main",
            #     "title": "Сальдо нач.",
            #     "type": "float",
            #     "col": 7,
            # },
            # {"name": "turn_credit_main", "title": "Кредит", "type": "float", "col": 8},
            # {
            #     "name": "end_debet_main",
            #     "title": "Сальдо кон.",
            #     "type": "float",
            #     "col": 9,
            # },
            #  {'name': 'turn_debet_proc', 'title': 'Процент дебет', 'type': 'float', 'col': 10},
            #  {'name': 'turn_credit_proc',
            #      'title': 'Процент кредит', 'type': 'float', 'col': 11},
            #  {'name': 'end_debet_proc',
            #      'title': 'Процент остаток', 'type': 'float', 'col': 12},
            {"name": "pdn", "title": "ПДН", "type": "float", "col": 13},
            #  {'name': 'period_common', 'title': 'Общий срок',
            #      'type': 'float', 'col': 14},
            #  {'name': 'date_finish', 'title': 'Крайняя дата',
            #      'type': 'date', 'col': 15},
            #  {'name': 'beg_debet_proc', 'title': 'Долг', 'type': 'float', 'col': 17},
            # {
            #     "name": "turn_debet_proc",
            #     "title": "Начисление",
            #     "type": "float",
            #     "col": 18,
            # },
            {
                "name": "turn_credit_proc",
                "title": "Оплата 1С",
                "type": "float",
                "col": 19,
            },
            {
                "name": "payment_archi",
                "title": "Оплата Archi",
                "type": "float",
                "col": 19,
            },
            # {"name": "report_date", "title": "Дата платежа", "type": "date", "col": 21},
            {
                "name": "end_debet_proc",
                "title": "Остаток платежа",
                "type": "float",
                "col": 20,
            },
            {
                "name": "report_frost",
                "title": "Дата заморозки",
                "type": "date",
                "col": 21,
            },
            {
                "name": "report_end_calculate",
                "title": "Дата начисления",
                "type": "date",
                "col": 22,
            },
            {
                "name": "count_days_last_month",
                "title": "К-во дней",
                "type": "int",
                "col": 23,
            },
            {
                "name": "turn_debet_proc_month",
                "title": "Начисление за месяц",
                "type": "float",
                "col": 24,
            },
            {
                "name": "count_days_delay",
                "title": "Просрочка",
                "type": "int",
                "col": 25,
            },
            {"name": "reserve", "title": "Разерв(проц)", "type": "int", "col": 26},
        ]

        num_format = "#,##0.00"
        self.workbook.write(0, 0, report.report_date, num_format_str=r"dd/mm/yyyy")
        row = 1
        col = 0
        self.workbook.write(row, col, "ФИО")
        for name in names:
            col += 1
            self.workbook.write(row, col, name["title"])
        col += 1
        self.workbook.write(row, col, "Резерв(осн.)")
        col += 1
        self.workbook.write(row, col, "Резерв(проц.)")
        row += 1
        curr_type = "Основной договор"
        col = 0
        for client in report.clients.values():
            for dog in client["dogovor"].values():
                if dog["type"] and dog["type"] != curr_type:
                    self.workbook.write(row, 0, dog["type"])
                    curr_type = dog["type"]
                    row += 1
                col = 0
                self.workbook.write(row, col, client["name"])
                client[
                    "name" + "_address"
                ] = f"{self.workbook.sheet.name}!{Utils.rowcol_to_cell(row,col)}"
                for name in names:
                    col += 1
                    try:
                        if dog.get(name["name"]):
                            value = (
                                float(dog[name["name"]])
                                if name["type"] == "float"
                                else (
                                    int(dog[name["name"]])
                                    if name["type"] == "int"
                                    else (
                                        to_date(dog[name["name"]])
                                        if name["type"] == "date"
                                        else (
                                            str(dog[name["name"]])
                                            if dog.get(name["name"])
                                            else None
                                        )
                                    )
                                )
                            )
                            if name["name"] == "report_frost":
                                if value != report.report_date and dog.get(
                                    "end_debet_proc"
                                ):
                                    self.workbook.write(
                                        row,
                                        col,
                                        value,
                                        num_format_str=r"dd/mm/yyyy"
                                        if name["type"] == "date"
                                        else None,
                                    )
                            elif name["name"] == "report_date":
                                if value != report.report_date:
                                    self.workbook.write(
                                        row,
                                        col,
                                        value,
                                        num_format_str=r"dd/mm/yyyy"
                                        if name["type"] == "date"
                                        else None,
                                    )
                            elif name["name"] == "count_days_delay":
                                x = (
                                    f"IF(ISBLANK({Utils.rowcol_to_cell(row,col-2,col_abs=True)}),"
                                    + f"0,"
                                    + f"{Utils.rowcol_to_cell(0,0,row_abs=True,col_abs=True)}-"
                                    + f"{Utils.rowcol_to_cell(row,2,col_abs=True)}-"
                                    + f"{Utils.rowcol_to_cell(row,6,col_abs=True)})"
                                )
                                f = f"IF(" + x + f'<=0,"",' + x + f")"
                                self.workbook.write(
                                    row,
                                    col,
                                    Formula(f),
                                    num_format_str=r"dd/mm/yyyy"
                                    if name["type"] == "date"
                                    else None,
                                )
                            else:
                                self.workbook.write(
                                    row,
                                    col,
                                    value,
                                    num_format_str=r"dd/mm/yyyy"
                                    if name["type"] == "date"
                                    else None,
                                )
                        elif name["name"] == "turn_debet_proc_month":
                            f = f"{Utils.rowcol_to_cell(row,3,col_abs=True)}*({Utils.rowcol_to_cell(row,4,col_abs=True)}/100)*{Utils.rowcol_to_cell(row,col-1,col_abs=True)}"
                            self.workbook.write(
                                row,
                                col,
                                Formula(f),
                                num_format_str=r"dd/mm/yyyy"
                                if name["type"] == "date"
                                else None,
                            )
                        elif name["name"] == "payment_archi":
                            if dog.get("payment"):
                                value = sum([x[2] for x in dog["payment"]])
                                self.workbook.write(
                                    row,
                                    col,
                                    value,
                                    num_format_str=r"dd/mm/yyyy"
                                    if name["type"] == "date"
                                    else None,
                                )
                        elif name["name"] == "reserve":
                            f = (
                                ""
                                + f'IF({Utils.rowcol_to_cell(row,col-1,col_abs=True)}="","",'
                                + f"IF(AND({Utils.rowcol_to_cell(row,col-1,col_abs=True)}<=7,{Utils.rowcol_to_cell(row,col-1,col_abs=True)}>=0),0,"
                                + f"IF({Utils.rowcol_to_cell(row,col-1,col_abs=True)}<=30,3/100,"
                                + f"IF({Utils.rowcol_to_cell(row,col-1,col_abs=True)}<=60,10/100,"
                                + f"IF({Utils.rowcol_to_cell(row,col-1,col_abs=True)}<=90,20/100,"
                                + f"IF({Utils.rowcol_to_cell(row,col-1,col_abs=True)}<=120,40/100,"
                                + f"IF({Utils.rowcol_to_cell(row,col-1,col_abs=True)}<=180,50/100,"
                                + f"IF({Utils.rowcol_to_cell(row,col-1,col_abs=True)}<=270,65/100,"
                                + f"IF({Utils.rowcol_to_cell(row,col-1,col_abs=True)}<=360,80/100,"
                                + f"99/100)))))))))"
                            )
                            self.workbook.write(row, col, Formula(f))
                    except Exception as ex:
                        print(
                            f"{self.workbook.sheet.name} ({name['name']}): {row}, {name['col']}, {value}"
                        )
                    dog[
                        name["name"] + "_address"
                    ] = f"{self.workbook.sheet.name}!{Utils.rowcol_to_cell(row,col)}"
                col += 1
                f = f'IF({Utils.rowcol_to_cell(row,col-1,col_abs=True)}="","",{Utils.rowcol_to_cell(row,col-9,col_abs=True)}*{Utils.rowcol_to_cell(row,col-1,col_abs=True)})'
                self.workbook.write(row, col, Formula(f), num_format_str=num_format)
                dog[
                    "rez_osn" + "_address"
                ] = f"{self.workbook.sheet.name}!{Utils.rowcol_to_cell(row,col)}"
                col += 1
                f = f'IF({Utils.rowcol_to_cell(row,col-2,col_abs=True)}="","",{Utils.rowcol_to_cell(row,col-5,col_abs=True)}*{Utils.rowcol_to_cell(row,col-1,col_abs=True)}/{Utils.rowcol_to_cell(row,col-10,col_abs=True)})'
                self.workbook.write(row, col, Formula(f), num_format_str=num_format)
                dog[
                    "rez_proc" + "_address"
                ] = f"{self.workbook.sheet.name}!{Utils.rowcol_to_cell(row,col)}"
                row += 1

    def write_result_weighted_average(self, result):
        if len(result) == 0:
            return
        names = [
            {"name": "stavka", "title": "Ставка"},
            {"name": "period", "title": "Срок"},
            {"name": "koef", "title": "Коэфф."},
        ]
        pattern_style = "pattern: pattern solid, fore_colour green; font: color yellow;"
        pattern_style_wa = (
            "pattern: pattern solid, fore_colour yellow; font: color black;"
        )
        num_format = "#,##0.00"
        index = 0
        for key, value in result.items():
            index += 1
            row = 0
            col = (index - 1) * 5 + 1
            if isinstance(value, dict):
                self.workbook.write(row, col - 1, key.split("_")[0])
                for name in names:
                    row += 1
                    self.workbook.write(row, col - 1, name["title"])
                    self.workbook.write(row, col, value[name["name"]])
                self.workbook.write(row + 1, col - 1, "Сумма")
                self.workbook.write(
                    row + 1,
                    col,
                    Formula(
                        f"SUM({Utils.rowcol_pair_to_cellrange(row+5,col+2,row+5+len(value['value'])-1,col+2)})"
                    ),
                    num_format_str=num_format,
                )
                self.workbook.write(row + 2, col - 1, "Сумма(ср.вз.)")
                self.workbook.write(
                    row + 2,
                    col,
                    Formula(
                        f"SUM({Utils.rowcol_pair_to_cellrange(row+5,col+3,row+5+len(value['value'])-1,col+3)})"
                    ),
                    num_format_str=num_format,
                )
                self.workbook.write(row + 3, col - 1, "Кол-во")
                self.workbook.write(
                    row + 3,
                    col,
                    Formula(
                        f"SUM({Utils.rowcol_pair_to_cellrange(row+5,col,row+5+len(value['value'])-1,col)})"
                    ),
                )
                sorted_value = sorted(value["value"].items(), key=lambda x: float(x[0]))
                row += 3
                row_start = row
                row += 1
                for val in sorted_value:
                    row += 1
                    self.workbook.write(row, col, float(val[1]))
                    self.workbook.write(row, col + 1, float(val[0]))
                    self.workbook.write(
                        row,
                        col + 2,
                        Formula(
                            f"{Utils.rowcol_to_cell(row,col)}*{Utils.rowcol_to_cell(row,col+1)}"
                        ),
                    )
                    self.workbook.write(
                        row,
                        col + 3,
                        Formula(
                            f"{Utils.rowcol_to_cell(row,col+2)}*{Utils.rowcol_to_cell(3,col)}"
                        ),
                    )
                self.workbook.write(
                    row + 1,
                    col + 2,
                    Formula(
                        f"SUM({Utils.rowcol_pair_to_cellrange(row_start+2,col+2,row,col+2)})"
                    ),
                    num_format_str=num_format,
                )
                self.workbook.write(
                    row + 1,
                    col + 3,
                    Formula(
                        f"SUM({Utils.rowcol_pair_to_cellrange(row_start+2,col+3,row,col+3)})"
                    ),
                    num_format_str=num_format,
                )
                row += 2
                for dog in value["parent"]:
                    self.workbook.write(
                        row,
                        col,
                        Formula(dog["name_address"])
                        if dog.get("name_address")
                        else dog.get("name", ""),
                    )
                    self.workbook.write(
                        row,
                        col + 1,
                        Formula(dog["number_address"])
                        if dog.get("number_address")
                        else dog.get("number", ""),
                    )
                    self.workbook.write(
                        row,
                        col + 2,
                        Formula(dog["summa_address"])
                        if dog.get("summa_address")
                        else dog.get("summa", ""),
                        num_format_str=num_format,
                    )
                    row += 1
        self.workbook.write(0, 2, "Общая сумма")
        self.workbook.write(
            0,
            3,
            Formula(f"SUM({Utils.rowcol_pair_to_cellrange(4,0,4,len(result)*5)})"),
            num_format_str=num_format,
        )
        self.workbook.write(1, 2, "Общая сумма(ср.вз.)")
        self.workbook.write(
            1,
            3,
            Formula(f"SUM({Utils.rowcol_pair_to_cellrange(5,0,5,len(result)*5)})"),
            num_format_str=num_format,
        )
        self.workbook.write(2, 2, "Сред.взвеш.")
        self.workbook.write(
            2,
            3,
            Formula(
                f"SUM({Utils.rowcol_pair_to_cellrange(3,0,3,len(result)*5)})/COUNT({Utils.rowcol_pair_to_cellrange(3,0,3,len(result)*5)})"
            ),
            style_string=pattern_style_wa,
        )

    def write_kategoria(self, kategoria):
        row = 0
        col = 0
        names = ["1", "2", "3", "4", "5", "6"]
        for name in names:
            self.workbook.write(row, col, name, "align: horiz center")
            col += 1
        row += 1
        col = 0
        nrow_start = len(kategoria.items()) + 3
        pattern_style_5 = (
            "pattern: pattern solid, fore_colour green; font: color yellow;"
        )
        num_format = "#,##0.00"
        pattern_style_3 = (
            "pattern: pattern solid, fore_colour orange; font: color white"
        )
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
                nrow_start += value["count4"] + 1
                row += 1
        self.workbook.write(row, col + 1, "Всего", "align: horiz left")
        if row > 7:
            self.workbook.write(
                row,
                col + 2,
                Formula(
                    f"SUM({Utils.rowcol_pair_to_cellrange(row-7,col+2,row-1,col+2)})"
                ),
                pattern_style_3,
                num_format,
            )
            self.workbook.write(
                row,
                col + 3,
                Formula(
                    f"SUM({Utils.rowcol_pair_to_cellrange(row-7,col+3,row-1,col+3)})"
                ),
                pattern_style_5,
                num_format,
            )
            self.workbook.write(
                row,
                col + 4,
                Formula(
                    f"SUM({Utils.rowcol_pair_to_cellrange(row-7,col+4,row-1,col+4)})"
                ),
                pattern_style_5,
                num_format,
            )
            self.workbook.write(
                row,
                col + 5,
                Formula(
                    f"SUM({Utils.rowcol_pair_to_cellrange(row-7,col+4,row-1,col+5)})"
                ),
                pattern_style_3,
                num_format,
            )

        row += 2
        self.workbook.write(row, col + 4, "(5)основная")
        self.workbook.write(row, col + 5, "(3,6)основная")
        self.workbook.write(row, col + 6, "(3,6)процент")
        self.workbook.write(row, col + 7, "Дней просрочки")
        for key, value in kategoria.items():
            row += 1
            self.workbook.write(row, col, key)
            for val in value["items"]:
                self.workbook.write(row, col + 1, val["name"])
                self.workbook.write(
                    row,
                    col + 2,
                    Formula(val["parent"]["number_address"])
                    if val["parent"].get("number_address")
                    else val["parent"].get("number"),
                )
                self.workbook.write(
                    row,
                    col + 3,
                    Formula(val["parent"]["pdn_address"])
                    if val["parent"].get("pdn_address")
                    else val["parent"].get("pdn"),
                )
                self.workbook.write(
                    row,
                    col + 4,
                    Formula(val["parent"]["turn_debet_main_address"])
                    if val["parent"].get("turn_debet_main_address")
                    else val["parent"].get("turn_debet_main"),
                )
                self.workbook.write(
                    row,
                    col + 5,
                    Formula(val["parent"]["end_debet_main_address"])
                    if val["parent"].get("end_debet_main_address")
                    else val["parent"].get("end_debet_main"),
                )
                self.workbook.write(
                    row,
                    col + 6,
                    Formula(val["parent"]["end_debet_proc_address"])
                    if val["parent"].get("end_debet_proc_address")
                    else val["parent"].get("end_debet_proc"),
                )
                self.workbook.write(
                    row,
                    col + 7,
                    Formula(val["parent"]["count_days_address"])
                    if val["parent"].get("count_days_address")
                    else val["parent"].get("count_days_delay"),
                )
                row += 1

    def write_reserve(self, report):
        row = 0
        col = 0
        names = [
            "Ставка",
            "Кол-во",
            "Основной",
            "Процент",
            "Резерв(осн)",
            "Резерв(проц)",
        ]
        for name in names:
            self.workbook.write(row, col, name, "align: horiz center")
            col += 2 if name == "Ставка" else 1
        row += 1
        col = 0
        nrow_start = len(report.reserve) + 3
        pattern_style_5 = (
            "pattern: pattern solid, fore_colour green; font: color yellow;"
        )
        num_format = "#,##0.00"
        pattern_style_3 = (
            "pattern: pattern solid, fore_colour orange; font: color white"
        )
        for value in report.reserve:
            self.workbook.write(row, col, value[1]["percent"])
            f = (
                f"COUNTIF({Utils.rowcol_pair_to_cellrange(nrow_start,col,nrow_start+len(report.clients),col)},"
                + f"{Utils.rowcol_to_cell(row,col,col_abs=True)}"
                + ")"
            )
            self.workbook.write(row, col + 2, Formula(f), pattern_style_3, num_format)
            f = (
                f"SUMIF({Utils.rowcol_pair_to_cellrange(nrow_start,col,nrow_start+len(report.clients),col)},"
                + f"{Utils.rowcol_to_cell(row,col,col_abs=True)},"
                + f"{Utils.rowcol_pair_to_cellrange(nrow_start,col+3,nrow_start+len(report.clients),col+3)}"
                + ")"
            )
            self.workbook.write(row, col + 3, Formula(f), pattern_style_3, num_format)
            f = (
                f"SUMIF({Utils.rowcol_pair_to_cellrange(nrow_start,col,nrow_start+len(report.clients),col)},"
                + f"{Utils.rowcol_to_cell(row,col,col_abs=True)},"
                + f"{Utils.rowcol_pair_to_cellrange(nrow_start,col+4,nrow_start+len(report.clients),col+4)}"
                + ")"
            )
            self.workbook.write(row, col + 4, Formula(f), pattern_style_3, num_format)
            f = (
                f"SUMIF({Utils.rowcol_pair_to_cellrange(nrow_start,col,nrow_start+len(report.clients),col)},"
                + f"{Utils.rowcol_to_cell(row,col,col_abs=True)},"
                + f"{Utils.rowcol_pair_to_cellrange(nrow_start,col+5,nrow_start+len(report.clients),col+5)}"
                + ")"
            )
            self.workbook.write(row, col + 5, Formula(f), pattern_style_5, num_format)
            f = (
                f"SUMIF({Utils.rowcol_pair_to_cellrange(nrow_start,col,nrow_start+len(report.clients),col)},"
                + f"{Utils.rowcol_to_cell(row,col,col_abs=True)},"
                + f"{Utils.rowcol_pair_to_cellrange(nrow_start,col+6,nrow_start+len(report.clients),col+6)}"
                + ")"
            )
            self.workbook.write(row, col + 6, Formula(f), pattern_style_5, num_format)
            row += 1

        row += 1
        col = 3
        for name in names[2:]:
            self.workbook.write(row, col, name, "align: horiz center")
            col += 1
        self.workbook.write(row, col, "Дней просрочки")
        row += 1
        col = 0
        nrow_start = 1
        for client in report.clients.values():
            for dog in client["dogovor"].values():
                self.workbook.write(row, col, Formula(dog["reserve_address"]))
                self.workbook.write(
                    row,
                    col + 1,
                    Formula(client["name_address"])
                    if client.get("name_address")
                    else client.get("name"),
                )
                self.workbook.write(
                    row,
                    col + 2,
                    Formula(dog["number_address"])
                    if dog.get("number_address")
                    else dog.get("number"),
                )
                self.workbook.write(
                    row,
                    col + 3,
                    Formula(dog["end_debet_main_address"])
                    if dog.get("end_debet_main_address")
                    else dog.get("end_debet_main"),
                    num_format_str=num_format,
                )
                self.workbook.write(
                    row,
                    col + 4,
                    Formula(dog["end_debet_proc_address"])
                    if dog.get("end_debet_proc_address")
                    else dog.get("end_debet_proc"),
                    num_format_str=num_format,
                )

                # f = f"{Utils.rowcol_to_cell(row,col+3,col_abs=True)}*{Utils.rowcol_to_cell(row,col,col_abs=True)}"
                f = dog["rez_osn_address"]
                self.workbook.write(row, col + 5, Formula(f), num_format_str=num_format)
                f = dog["rez_proc_address"]
                # f = f"{Utils.rowcol_to_cell(row,col+4,col_abs=True)}*({Utils.rowcol_to_cell(row,col+5,col_abs=True)}/{Utils.rowcol_to_cell(row,col+3,col_abs=True)})"
                self.workbook.write(row, col + 6, Formula(f), num_format_str=num_format)
                self.workbook.write(
                    row,
                    col + 7,
                    Formula(dog["count_days_address"])
                    if dog.get("count_days_address")
                    else dog.get("count_days_delay"),
                    num_format_str=num_format,
                )
            nrow_start += 1
            row += 1

    def write_payment(self, report):
        row = 0
        col = 0
        self.workbook.write(row, 0, "Название")
        self.workbook.write(row, 1, "Договор")
        self.workbook.write(row, 2, "Сумма в 1С")
        self.workbook.write(row, 3, "Сумма в Archi")
        self.workbook.write(row, 4, "Дата")

        for client in report.clients.values():
            num_format = "#,##0.00"
            for dog in client["dogovor"].values():
                if dog.get("payment"):
                    sum_on_archi_str = ""
                    sum_on_1c_str = ""
                    sum_on_archi = 0
                    sum_on_1c = 0
                    d = None
                    for pay in dog["payment"]:
                        if pay[3].date() < report.report_date and pay[5] == 1:
                            sum_on_archi += pay[2]
                            sum_on_archi_str += f"+{pay[2]}"
                            d = pay[3].date()
                    if sum_on_archi > 0 and dog.get("plat"):
                        for pay in dog["plat"]:
                            if (
                                pay.get("turn_credit_proc")
                                and datetime.datetime.strptime(
                                    pay["date_proc"], "%d.%m.%y"
                                ).date()
                                < report.report_date
                            ):
                                sum_on_1c += float(pay["turn_credit_proc"])
                                sum_on_1c_str += f"+{float(pay['turn_credit_proc'])}"
                        if float(sum_on_archi) != sum_on_1c:
                            self.workbook.write(
                                row,
                                0,
                                Formula(client["name_address"])
                                if client.get("name_address")
                                else client.get("name"),
                            )
                            self.workbook.write(
                                row,
                                1,
                                Formula(dog["number_address"])
                                if dog.get("number_address")
                                else dog.get("number_address"),
                            )
                            try:
                                self.workbook.write(
                                    row,
                                    2,
                                    Formula(sum_on_1c_str.strip("+"))
                                    if sum_on_1c_str
                                    else 0,
                                    num_format_str=num_format,
                                )
                                self.workbook.write(
                                    row,
                                    3,
                                    Formula(sum_on_archi_str.strip("+"))
                                    if sum_on_archi_str
                                    else 0,
                                    num_format_str=num_format,
                                )
                            except Exception as ex:
                                pass
                            self.workbook.write(row, 4, d, num_format_str="dd/mm/yyyy")
                            row += 1


# from xlwt import Utils
# print Utils.rowcol_pair_to_cellrange(2,2,12,2)
# print Utils.rowcol_to_cell(13,2)
#  ws.write(i, 2, xlwt.Formula("$A$%d+$B$%d" % (i+1, i+1)))
