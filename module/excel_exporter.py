import datetime
from xlwt import Utils, Formula, XFStyle
from module.file_readers import get_file_write
from module.helpers import to_date, get_value_attr, get_max_margin_rate
from module.data import *


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
        # self.workbook.addSheet("Платежи")
        # self.write_payment(report)
        # self.workbook.addSheet("error")
        # self.write_errors(report.warnings)
        return self.workbook.save()

    def write_errors(self, errors):
        row = 0
        col = 0
        for item in errors:
            self.workbook.write(row, col, item)
            row += 1

    def write_clients(self, report) -> bool:
        def get_names():
            return [
                {"name": "number", "title": "Номер", "type": ""},
                {"name": "date", "title": "Дата", "type": "date"},
                {"name": "summa", "title": "Сумма", "type": "float"},
                {"name": "rate", "title": "Ставка", "type": "float"},
                {"name": "tarif", "title": "Тариф", "type": ""},
                {"name": "count_days", "title": "Срок", "type": "float"},
                {"name": "pdn", "title": "ПДН", "type": "float"},
                {"name": "debet_beg_main", "title": "Д(58)", "type": "float"},
                {"name": "credit_beg_main", "title": "К(58)", "type": "float"},
                {"name": "debet_main", "title": "Д(58)", "type": "float"},
                {"name": "credit_main", "title": "К(58)", "type": "float"},
                {"name": "debet_end_main", "title": "Д(58)", "type": "float"},
                {"name": "credit_end_main", "title": "К(58)", "type": "float"},
                {"name": "debet_beg_proc", "title": "Д(76)", "type": "float"},
                {"name": "credit_beg_proc", "title": "К(76)", "type": "float"},
                {"name": "debet_proc", "title": "Д(76)", "type": "float"},
                {"name": "credit_proc", "title": "К(76)", "type": "float"},
                {"name": "debet_end_proc", "title": "Д(76)", "type": "float"},
                {"name": "credit_end_proc", "title": "К(76)", "type": "float"},
                {"name": "credit_main", "title": "1С(осн)", "type": "float"},
                {
                    "name": "credit_proc",
                    "title": "1С(проц)",
                    "type": "float",
                },
                {"name": "payments_base", "title": "Archi", "type": "float"},
                {"name": "date_frozen", "title": "Дата заморозки", "type": "date"},
                {
                    "name": "count_days_common",
                    "title": "Дней\n(всего)",
                    "type": "int",
                },
                {"name": "calculate_percent", "title": "Проц.всего", "type": "float"},
                {
                    "name": "count_days_period",
                    "title": "Дней\n(месяц)",
                    "type": "int",
                },
                {
                    "name": "calculate_period",
                    "title": "Проц.месяц",
                    "type": "float",
                },
                {"name": "calc_debet_end_main", "title": "Остаток(осн)", "type": "float"},
                {"name": "calc_debet_end_proc", "title": "Остаток(проц)", "type": "float"},
                {"name": "count_days_delay", "title": "Просрочка", "type": "int"},
                {"name": "reserve", "title": "Разерв(%)", "type": "int", "col": 26},
            ]

        def write_value_attribute(value):
            nonlocal name, row, order
            try:
                if name["name"] == "date_calculate":
                    if value != report.report_date:
                        self.workbook.write(
                            row, name["col"], value, type_name=name["type"]
                        )
                elif name["name"] == "tarif":
                    self.workbook.write(
                        row, name["col"], value.code, type_name=name["type"]
                    )
                elif name["name"] == "count_days_common":
                    f = f'IF({Utils.rowcol_to_cell(row,get_col("date_frozen"),col_abs=True)}="",'
                    f += f'{Utils.rowcol_to_cell(0,0,col_abs=True)},{Utils.rowcol_to_cell(row,get_col("date_frozen"),col_abs=True)})-'
                    f += f'{Utils.rowcol_to_cell(row,get_col("date"),col_abs=True)}'
                    self.workbook.write(
                        row, name["col"], Formula(f), type_name=name["type"]
                    )
                elif name["name"] == "count_days_delay":
                    f = f'IF({Utils.rowcol_to_cell(row,get_col("count_days_common"),col_abs=True)}-'
                    f += f'{Utils.rowcol_to_cell(row,get_col("count_days"),col_abs=True)}>0,'
                    f += f'{Utils.rowcol_to_cell(row,get_col("count_days_common"),col_abs=True)}-'
                    f += f'{Utils.rowcol_to_cell(row,get_col("count_days"),col_abs=True)},'
                    f += f'"")'
                    self.workbook.write(
                        row, name["col"], Formula(f), type_name=name["type"]
                    )
                elif name["name"] == "payments_1c":
                    value = sum(
                        [
                            x.summa
                            for x in order.payments_1c
                            if x.type == "O" and x.category == "C" and x.kind == "proc"
                        ]
                    )
                    self.workbook.write(
                        row,
                        name["col"],
                        value if value > 0 else "",
                        type_name=name["type"],
                    )
                elif name["name"] == "payments_base":
                    if order.payments_base:
                        value = sum([x[2] for x in value])
                        self.workbook.write(
                            row, name["col"], value, type_name=name["type"]
                        )
                else:
                    self.workbook.write(row, name["col"], value, type_name=name["type"])
            except Exception as ex:
                print(
                    f"{self.workbook.sheet.name} ({name['name']}): {row}, {name['col']}, {value}"
                )

        def write_function():
            nonlocal name, row, order
            try:
                value = ""
                if name["name"] == "calculate_period":
                    f = f"{Utils.rowcol_to_cell(row,get_col('summa'),col_abs=True)}*"
                    f += f"({Utils.rowcol_to_cell(row,get_col('rate'),col_abs=True)}/100)*"
                    f += f"{Utils.rowcol_to_cell(row,get_col('count_days_period'),col_abs=True)}"
                    self.workbook.write(
                        row, name["col"], Formula(f), type_name=name["type"]
                    )
                elif name["name"] == "calculate_percent":
                    summa_max = order.summa * Decimal(get_max_margin_rate(order.date))
                    f = f"MAX(MIN({summa_max},"
                    f += f"{Utils.rowcol_to_cell(row,get_col('summa'),col_abs=True)}*"
                    f += f"({Utils.rowcol_to_cell(row,get_col('rate'),col_abs=True)}/100)*"
                    f += f"{Utils.rowcol_to_cell(row,get_col('count_days_common'),col_abs=True)})-"
                    f += f"{Utils.rowcol_to_cell(row,get_col('credit_main'),col_abs=True)},0)"
                    self.workbook.write(
                        row, name["col"], Formula(f), type_name=name["type"]
                    )
                elif name["name"] == "calc_debet_end_main":
                    if order.debet_main == 0:
                        f = f"MAX({Utils.rowcol_to_cell(row,get_col('summa'),col_abs=True)}-"
                        f += f"{Utils.rowcol_to_cell(row,get_col('credit_main'),col_abs=True)},0)"
                        self.workbook.write(
                            row, name["col"], Formula(f), type_name=name["type"]
                        )
                    else:
                        self.workbook.write(
                            row, name["col"], order.debet_main, type_name=name["type"]
                        )
                elif name["name"] == "calc_debet_end_proc":
                        f = f"MAX({Utils.rowcol_to_cell(row,get_col('calculate_percent'),col_abs=True)}-"
                        f += f"{Utils.rowcol_to_cell(row,get_col('credit_proc'),col_abs=True)},0)"
                        self.workbook.write(
                            row, name["col"], Formula(f), type_name=name["type"]
                        )
                        summa_max = order.summa * Decimal(
                            get_max_margin_rate(order.date)
                        )
                elif name["name"] == "reserve":
                    col = get_col("count_days_delay")
                    f = (
                        ""
                        + f'IF({Utils.rowcol_to_cell(row,col,col_abs=True)}="","",'
                        + f"IF(AND({Utils.rowcol_to_cell(row,col,col_abs=True)}<=7,{Utils.rowcol_to_cell(row,col-1,col_abs=True)}>=0),0,"
                        + f"IF({Utils.rowcol_to_cell(row,col,col_abs=True)}<=30,3/100,"
                        + f"IF({Utils.rowcol_to_cell(row,col,col_abs=True)}<=60,10/100,"
                        + f"IF({Utils.rowcol_to_cell(row,col,col_abs=True)}<=90,20/100,"
                        + f"IF({Utils.rowcol_to_cell(row,col,col_abs=True)}<=120,40/100,"
                        + f"IF({Utils.rowcol_to_cell(row,col,col_abs=True)}<=180,50/100,"
                        + f"IF({Utils.rowcol_to_cell(row,col,col_abs=True)}<=270,65/100,"
                        + f"IF({Utils.rowcol_to_cell(row,col,col_abs=True)}<=360,80/100,"
                        + f"99/100)))))))))"
                    )
                    self.workbook.write(row, name["col"], Formula(f))
            except Exception as ex:
                print(
                    f"{self.workbook.sheet.name} ({name['name']}): {row}, {name['col']}, {value}"
                )

        def get_col(name: str) -> int:
            nonlocal names
            s = [x["col"] for x in names if x["name"] == name]
            return s[0] if s else 0

        def calculate_rezerves_main():
            f = f'IF({Utils.rowcol_to_cell(row,get_col("reserve"),col_abs=True)}="","",'
            f += f'{Utils.rowcol_to_cell(row,get_col("calc_debet_end_main"),col_abs=True)}*'
            f += f'{Utils.rowcol_to_cell(row,get_col("reserve"),col_abs=True)})'
            self.workbook.write(
                row, len(names) + 1, Formula(f), num_format_str=num_format
            )
            order.link[
                "reserve_main" + "_address"
            ] = f"{self.workbook.sheet.name}!{Utils.rowcol_to_cell(row,len(names)+1)}"

        def calculate_rezerves_proc():
            f = f'IF({Utils.rowcol_to_cell(row,get_col("reserve"),col_abs=True)}="","",'
            f += f'{Utils.rowcol_to_cell(row,get_col("debet_end_proc"),col_abs=True)}*'
            f += f'{Utils.rowcol_to_cell(row,get_col("reserve"),col_abs=True)})'
            self.workbook.write(
                row, len(names) + 2, Formula(f), num_format_str=num_format
            )
            order.link[
                "reserve_proc" + "_address"
            ] = f"{self.workbook.sheet.name}!{Utils.rowcol_to_cell(row,len(names) + 2)}"

        def write_header():
            row = 1
            self.workbook.write(row, 0, "ФИО")
            for col, name in enumerate(names, 1):
                name["col"] = col
                self.workbook.write(row, col, name["title"])
            self.workbook.write(row, len(names) + 1, "Резерв(осн.)")
            self.workbook.write(row, len(names) + 2, "Резерв(проц.)")

        # ----------------------------------------------------------------------------------------------------------
        names = get_names()
        num_format = "#,##0.00"
        self.workbook.write(0, 0, report.report_date, num_format_str=r"dd/mm/yyyy")
        write_header()
        row = 2
        curr_type = "Основной договор"
        client: Client = Client()
        order: Order = Order()
        for client in report.clients.values():
            for order in client.orders:
                if order.type and order.type != curr_type:
                    self.workbook.write(row, 0, order.type)
                    curr_type = order.type
                    row += 1

                self.workbook.write(row, 0, client.name)
                client.link[
                    "name" + "_address"
                ] = f"{self.workbook.sheet.name}!{Utils.rowcol_to_cell(row,0)}"
                for name in names:
                    value = (
                        getattr(order, name["name"])
                        if hasattr(order, name["name"])
                        else None
                    )
                    if value:
                        write_value_attribute(value)
                    else:
                        write_function()
                    order.link[
                        name["name"] + "_address"
                    ] = f"{self.workbook.sheet.name}!{Utils.rowcol_to_cell(row,name['col'])}"
                calculate_rezerves_main()
                calculate_rezerves_proc()
                row += 1

    # %% Средневзвешенная
    def write_result_weighted_average(self, result: dict):
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
                order: Order = Order()
                for order in value["parent"]:
                    self.workbook.write(
                        row,
                        col,
                        Formula(order.link["name_address"])
                        if order.link.get("name_address")
                        else order.client.name,
                    )
                    self.workbook.write(
                        row,
                        col + 1,
                        Formula(order.link["number_address"])
                        if order.link.get("number_address")
                        else order.number,
                    )
                    self.workbook.write(
                        row,
                        col + 2,
                        Formula(order.link["summa_address"])
                        if order.link.get("summa_address")
                        else order.summa,
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

    # %% Категория
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
                row += 1

    # %% Резервы
    def write_reserve(self, report: dict):
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
            reserve: Reserve = Reserve()
            reserve = value[1]
            self.workbook.write(row, col, reserve.percent)
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
        client: Client = Client()
        for client in report.clients.values():
            for order in client.orders:
                self.workbook.write(
                    row, col, Formula(order.link.get("reserve_address", ""))
                )
                self.workbook.write(
                    row,
                    col + 1,
                    Formula(client.link.get("name_address", ""))
                    if client.link.get("name_address")
                    else client.name,
                )
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
                    Formula(order.link["debet_end_main_address"])
                    if order.link.get("debet_end_main_address")
                    else order.debet_main,
                    num_format_str=num_format,
                )
                self.workbook.write(
                    row,
                    col + 4,
                    Formula(order.link["debet_end_proc_address"])
                    if order.link.get("debet_end_proc_address")
                    else order.debet_proc,
                    num_format_str=num_format,
                )

                f = order.link.get("reserve_main_address", "")
                self.workbook.write(row, col + 5, Formula(f), num_format_str=num_format)
                f = order.link.get("reserve_proc_address")
                self.workbook.write(row, col + 6, Formula(f), num_format_str=num_format)
                self.workbook.write(
                    row,
                    col + 7,
                    Formula(order.link.get("count_days_delay_address", ""))
                    if order.link.get("count_days_delay_address")
                    else order.count_days_delay,
                    num_format_str=num_format,
                )
                nrow_start += 1
                row += 1

    # %% Платежи
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
