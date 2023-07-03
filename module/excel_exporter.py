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
        self.workbook.addSheet("ОперацииВручную")
        self.write_payment(report)
        # self.workbook.addSheet("63")
        # self.write_payment(report, "proc")
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
            return (
                [
                    {"name": "number", "title": "Номер", "type": ""},
                    {"name": "date", "title": "Дата", "type": "date"},
                    {"name": "summa", "title": "Сумма", "type": "float"},
                    {"name": "rate", "title": "Ставка", "type": "float"},
                    {"name": "tarif", "title": "Тариф", "type": ""},
                    {"name": "count_days", "title": "Срок", "type": "float"},
                    {"name": "pdn", "title": "ПДН", "type": "float"},
                    {"name": "credit_main",
                        "title": "оплата 1С(осн)", "type": "float"},
                    {
                        "name": "credit_proc",
                        "title": "оплата 1С(проц)",
                        "type": "float",
                    },
                ]
                + ([
                    {"name": "payments_base", "title": "Archi", "type": "float"},
                    {"name": "date_frozen", "title": "Дата заморозки", "type": "date"},
                ]
                    if report.is_archi
                    else [])
                + [
                    {
                        "name": "count_days_common",
                        "title": "Дней\n(всего)",
                        "type": "int",
                    },
                    {
                        "name": "calculate_percent",
                        "title": "Проц.всего",
                        "type": "float",
                    },
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
                    {
                        "name": "calc_debet_end_main",
                        "title": "Остаток(осн)",
                        "type": "float",
                    },
                    {
                        "name": "debet_end_proc",  # "calc_debet_end_proc",
                        "title": "Остаток(проц)",
                        "type": "float",
                    },
                    {"name": "count_days_delay", "title": "Просрочка", "type": "int"},
                    {"name": "reserve_percent",
                        "title": "Разерв(%)", "type": "int"},
                    {
                        "name": "calc_reserve_main",
                        "title": "Разерв(осн)",
                        "type": "float",
                    },
                    {"name": "credit_end_main",
                        "title": "Кредит(59)", "type": "float"},
                    {
                        "name": "calc_reserve_proc",
                        "title": "Разерв(проц)",
                        "type": "float",
                    },
                    {"name": "credit_end_proc",
                        "title": "Кредит(63)", "type": "float"},
                    # {"name": "debet_beg_main", "title": "СальдНач(Д58)", "type": "float"},
                    # {"name": "credit_beg_main", "title": "СальдоНач(К58)", "type": "float"},
                    # {"name": "debet_main", "title": "Оборот(Д58)", "type": "float"},
                    # {"name": "credit_main", "title": "Оборот(К58)", "type": "float"},
                    # {"name": "debet_end_main", "title": "СальдКон(Д58)", "type": "float"},
                    # {"name": "credit_end_main", "title": "СальдКон(К58)", "type": "float"},
                    # {"name": "debet_beg_proc", "title": "СальдНач(Д76)", "type": "float"},
                    # {"name": "credit_beg_proc", "title": "СальдНач(К76)", "type": "float"},
                    # {"name": "debet_proc", "title": "Оборот(Д76)", "type": "float"},
                    # {"name": "credit_proc", "title": "Оборот(К76)", "type": "float"},
                    # {"name": "debet_end_proc", "title": "СальдКон(Д76)", "type": "float"},
                    # {"name": "credit_end_proc", "title": "СальдКон(К76)", "type": "float"},
                ]
            )

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
                    if value:
                        value = sum([x[2] for x in value])
                        self.workbook.write(
                            row, name["col"], value, type_name=name["type"]
                        )
                else:
                    self.workbook.write(
                        row, name["col"], value, type_name=name["type"])
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
                    summa_max = order.summa * \
                        Decimal(get_max_margin_rate(order.date))
                    f = f"MAX(MIN({summa_max},"
                    f += f"{Utils.rowcol_to_cell(row,get_col('summa'),col_abs=True)}*"
                    f += f"({Utils.rowcol_to_cell(row,get_col('rate'),col_abs=True)}/100)*"
                    f += f"{Utils.rowcol_to_cell(row,get_col('count_days_common'),col_abs=True)})-"
                    f += f"{Utils.rowcol_to_cell(row,get_col('credit_main'),col_abs=True)},0)"
                    self.workbook.write(
                        row, name["col"], Formula(f), type_name=name["type"]
                    )
                    order.calculate_percent = max(
                        min(
                            summa_max,
                            order.summa
                            * Decimal(order.rate)
                            / 100
                            * Decimal(order.count_days_common),
                        )
                        - order.credit_end_main,
                        0,
                    )
                elif name["name"] == "calc_debet_end_main":
                    f = f"MAX({Utils.rowcol_to_cell(row,get_col('summa'),col_abs=True)}-"
                    f += f"{Utils.rowcol_to_cell(row,get_col('credit_main'),col_abs=True)},0)"
                    self.workbook.write(
                        row, name["col"], Formula(f), type_name=name["type"]
                    )
                    order.debet_end_main = max(
                        order.summa - order.credit_main, 0)
                elif name["name"] == "calc_debet_end_proc":
                    f = f"MAX({Utils.rowcol_to_cell(row,get_col('calculate_percent'),col_abs=True)}-"
                    f += f"{Utils.rowcol_to_cell(row,get_col('credit_proc'),col_abs=True)},0)"
                    self.workbook.write(
                        row, name["col"], Formula(f), type_name=name["type"]
                    )
                    order.debet_end_proc = max(
                        order.calculate_percent - order.credit_proc, 0
                    )
                elif name["name"] == "reserve_percent":
                    col = get_col("count_days_delay")
                    col1 = get_col("pdn")
                    f = (
                        ""
                        + f'IF(AND({Utils.rowcol_to_cell(row,col,col_abs=True)}="",{Utils.rowcol_to_cell(row,col1,col_abs=True)}=""),"",'
                        + f'IF({Utils.rowcol_to_cell(row,col1,col_abs=True)}="","",'
                        + f'IF(AND({Utils.rowcol_to_cell(row,col,col_abs=True)}="",'
                        + f'{Utils.rowcol_to_cell(row,get_col("calc_debet_end_main"),col_abs=True)}>=10000,'
                        + f'{Utils.rowcol_to_cell(row,get_col("pdn"),col_abs=True)}>=0.5),0,'
                        + f'IF(AND({Utils.rowcol_to_cell(row,col,col_abs=True)}="",'
                        + f'OR({Utils.rowcol_to_cell(row,get_col("calc_debet_end_main"),col_abs=True)}<10000,'
                        + f'{Utils.rowcol_to_cell(row,get_col("pdn"),col_abs=True)}<0.5)),-1,'
                        + f"IF(AND({Utils.rowcol_to_cell(row,col,col_abs=True)}<=7,{Utils.rowcol_to_cell(row,col-1,col_abs=True)}>=0),0,"
                        + f"IF({Utils.rowcol_to_cell(row,col,col_abs=True)}<=30,3/100,"
                        + f"IF({Utils.rowcol_to_cell(row,col,col_abs=True)}<=60,10/100,"
                        + f"IF({Utils.rowcol_to_cell(row,col,col_abs=True)}<=90,20/100,"
                        + f"IF({Utils.rowcol_to_cell(row,col,col_abs=True)}<=120,40/100,"
                        + f"IF({Utils.rowcol_to_cell(row,col,col_abs=True)}<=180,50/100,"
                        + f"IF({Utils.rowcol_to_cell(row,col,col_abs=True)}<=270,65/100,"
                        + f"IF({Utils.rowcol_to_cell(row,col,col_abs=True)}<=360,80/100,"
                        + f"IF({Utils.rowcol_to_cell(row,col,col_abs=True)}>360,99/100,"
                        + f'IF({Utils.rowcol_to_cell(row,get_col("count_days"),col_abs=True)}>31,0,'
                        + f'""))))))))))))))'

                    )
                    self.workbook.write(row, name["col"], Formula(f))
                elif name["name"] == "count_days_common":
                    if get_col("date_frozen") != 0:
                        f = f'IF({Utils.rowcol_to_cell(row,get_col("date_frozen"),col_abs=True)}="",'
                        f += f"{Utils.rowcol_to_cell(0,0,col_abs=True,row_abs=True)},"
                        f += f'{Utils.rowcol_to_cell(row,get_col("date_frozen"),col_abs=True)})-'
                        f += f'{Utils.rowcol_to_cell(row,get_col("date"),col_abs=True)}'
                    else:
                        f = f'{Utils.rowcol_to_cell(0,0,col_abs=True,row_abs=True)}-{Utils.rowcol_to_cell(row,get_col("date"),col_abs=True)}'
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
                elif name["name"] == "calc_reserve_main":
                    calculate_rezerves_main()
                elif name["name"] == "calc_reserve_proc":
                    calculate_rezerves_proc()
            except Exception as ex:
                print(
                    f"{self.workbook.sheet.name} ({name['name']}): {row}, {name['col']}, {value}"
                )

        def get_col(name: str) -> int:
            nonlocal names
            s = [x["col"] for x in names if x["name"] == name]
            return s[0] if s else 0

        def calculate_rezerves_main():
            f = f'IF({Utils.rowcol_to_cell(row,get_col("reserve_percent"),col_abs=True)}="","",'
            f += f'IF({Utils.rowcol_to_cell(row,get_col("reserve_percent"),col_abs=True)}<0,"",'
            f += f'IF({Utils.rowcol_to_cell(row,get_col("reserve_percent"),col_abs=True)}=0,'
            f += f'{Utils.rowcol_to_cell(row,get_col("calc_debet_end_main"),col_abs=True)}*1/10,'
            f += f'{Utils.rowcol_to_cell(row,get_col("calc_debet_end_main"),col_abs=True)}*'
            f += f'{Utils.rowcol_to_cell(row,get_col("reserve_percent"),col_abs=True)}'
            f += f")))"
            self.workbook.write(row, name["col"], Formula(
                f), num_format_str=num_format)
            order.link[
                "calc_reserve_main" + "_address"
            ] = f"{self.workbook.sheet.name}!{Utils.rowcol_to_cell(row,len(names)+1)}"
            order.calc_reserve_main = order.debet_end_main * \
                Decimal(order.percent)

        def calculate_rezerves_proc():
            f = f'IF({Utils.rowcol_to_cell(row,get_col("reserve_percent"),col_abs=True)}="","",'
            f += f'IF({Utils.rowcol_to_cell(row,get_col("reserve_percent"),col_abs=True)}<0,"",'
            f += f'IF({Utils.rowcol_to_cell(row,get_col("reserve_percent"),col_abs=True)}=0,'
            f += f'{Utils.rowcol_to_cell(row,get_col("debet_end_proc"),col_abs=True)}*1/10,'
            f += f'{Utils.rowcol_to_cell(row,get_col("debet_end_proc"),col_abs=True)}*'
            f += f'{Utils.rowcol_to_cell(row,get_col("reserve_percent"),col_abs=True)}'
            f += f")))"
            self.workbook.write(row, name["col"], Formula(
                f), num_format_str=num_format)
            order.link[
                "calc_reserve_proc" + "_address"
            ] = f"{self.workbook.sheet.name}!{Utils.rowcol_to_cell(row,len(names) + 2)}"
            order.calc_reserve_proc = order.debet_end_proc * \
                Decimal(order.percent)

        def write_header():
            row = 1
            self.workbook.write(row, 0, "ФИО")
            for col, name in enumerate(names, 1):
                name["col"] = col
                self.workbook.write(row, col, name["title"])
            # self.workbook.write(row, len(names) + 1, "Резерв(осн.)")
            # self.workbook.write(row, len(names) + 2, "Резерв(проц.)")

        # ----------------------------------------------------------------------------------------------------------
        names = get_names()
        num_format = "#,##0.00"
        self.workbook.write(0, 0, report.report_date,
                            num_format_str=r"dd/mm/yyyy")
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
                sorted_value = sorted(
                    value["value"].items(), key=lambda x: float(x[0]))
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
            Formula(
                f"SUM({Utils.rowcol_pair_to_cellrange(4,0,4,len(result)*5)})"),
            num_format_str=num_format,
        )
        self.workbook.write(1, 2, "Общая сумма(ср.вз.)")
        self.workbook.write(
            1,
            3,
            Formula(
                f"SUM({Utils.rowcol_pair_to_cellrange(5,0,5,len(result)*5)})"),
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
                    Formula(order.link["calc_debet_end_main_address"])
                    if order.link.get("calc_debet_end_main_address")
                    else order.debet_main,
                )
                self.workbook.write(
                    row,
                    col + 6,
                    Formula(order.link["calc_debet_end_proc_address"])
                    if order.link.get("calc_debet_end_proc_address")
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
        def fill_table(nrow_start: int, row: int, col: int):
            nonlocal pattern_style_3, num_format
            f = (
                f"COUNTIF({Utils.rowcol_pair_to_cellrange(nrow_start,col,nrow_start+len(report.clients),col)},"
                + f"{Utils.rowcol_to_cell(row,col,col_abs=True)}"
                + ")"
            )
            self.workbook.write(row, col + 2, Formula(f),
                                pattern_style_3, num_format)
            f = (
                f"SUMIF({Utils.rowcol_pair_to_cellrange(nrow_start,col,nrow_start+len(report.clients),col)},"
                + f"{Utils.rowcol_to_cell(row,col,col_abs=True)},"
                + f"{Utils.rowcol_pair_to_cellrange(nrow_start,col+3,nrow_start+len(report.clients),col+3)}"
                + ")"
            )
            self.workbook.write(row, col + 3, Formula(f),
                                pattern_style_3, num_format)
            f = (
                f"SUMIF({Utils.rowcol_pair_to_cellrange(nrow_start,col,nrow_start+len(report.clients),col)},"
                + f"{Utils.rowcol_to_cell(row,col,col_abs=True)},"
                + f"{Utils.rowcol_pair_to_cellrange(nrow_start,col+4,nrow_start+len(report.clients),col+4)}"
                + ")"
            )
            self.workbook.write(row, col + 4, Formula(f),
                                pattern_style_3, num_format)
            f = (
                f"SUMIF({Utils.rowcol_pair_to_cellrange(nrow_start,col,nrow_start+len(report.clients),col)},"
                + f"{Utils.rowcol_to_cell(row,col,col_abs=True)},"
                + f"{Utils.rowcol_pair_to_cellrange(nrow_start,col+5,nrow_start+len(report.clients),col+5)}"
                + ")"
            )
            self.workbook.write(row, col + 5, Formula(f),
                                pattern_style_5, num_format)
            f = (
                f"SUMIF({Utils.rowcol_pair_to_cellrange(nrow_start,col,nrow_start+len(report.clients),col)},"
                + f"{Utils.rowcol_to_cell(row,col,col_abs=True)},"
                + f"{Utils.rowcol_pair_to_cellrange(nrow_start,col+6,nrow_start+len(report.clients),col+6)}"
                + ")"
            )
            self.workbook.write(row, col + 6, Formula(f),
                                pattern_style_5, num_format)
            return

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
            reserve: Reserve = value[1]
            self.workbook.write(row, col, reserve.percent)
            fill_table(nrow_start, row, col)
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
                    row, col, Formula(order.link.get(
                        "reserve_percent_address", ""))
                )
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
                    Formula(order.link["calc_debet_end_main_address"])
                    if order.link.get("calc_debet_end_main_address")
                    else order.debet_main,
                    num_format_str=num_format,
                )
                self.workbook.write(
                    row,
                    col + 4,
                    Formula(order.link["calc_debet_end_proc_address"])
                    if order.link.get("calc_debet_end_proc_address")
                    else order.debet_proc,
                    num_format_str=num_format,
                )

                f = order.link.get("calc_reserve_main_address", "")
                self.workbook.write(row, col + 5, Formula(f),
                                    num_format_str=num_format)
                f = order.link.get("calc_reserve_proc_address")
                self.workbook.write(row, col + 6, Formula(f),
                                    num_format_str=num_format)
                m = order.link.get("count_days_delay_address", "")
                f = f'IF({m}=0,"",{m})'
                self.workbook.write(
                    row,
                    col + 7,
                    Formula(f)
                    if order.link.get("count_days_delay_address")
                    else order.count_days_delay,
                    num_format_str=num_format,
                )
                nrow_start += 1
                row += 1

    # %% Платежи
    def write_payment(self, report: dict):
        def __write(sub: str):
            nonlocal client, order, row, col, num_format
            summa_1c = getattr(order, f"credit_end_{sub}")
            summa_calc = getattr(order, f"calc_reserve_{sub}")
            # if (summa_1c != 0 or summa_calc != 0) and summa_1c != summa_calc:
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
                Formula(order.link[f"credit_end_{sub}_address"]),
                num_format_str=num_format,
            )
            self.workbook.write(
                row,
                col + 3,
                Formula(order.link[f"calc_reserve_{sub}_address"]),
                num_format_str=num_format,
            )
            f = f"IF("
            f += f'IF({Utils.rowcol_to_cell(row,col+2,col_abs=True)}="",0,{Utils.rowcol_to_cell(row,col+2,col_abs=True)})>'
            f += f'IF({Utils.rowcol_to_cell(row,col+3,col_abs=True)}="",0,{Utils.rowcol_to_cell(row,col+3,col_abs=True)}),'
            f += f'IF({Utils.rowcol_to_cell(row,col+2,col_abs=True)}="",0,{Utils.rowcol_to_cell(row,col+2,col_abs=True)})-'
            f += f'IF({Utils.rowcol_to_cell(row,col+3,col_abs=True)}="",0,{Utils.rowcol_to_cell(row,col+3,col_abs=True)}),'
            f += f'IF({Utils.rowcol_to_cell(row,col+3,col_abs=True)}="",0,{Utils.rowcol_to_cell(row,col+3,col_abs=True)})-'
            f += f'IF({Utils.rowcol_to_cell(row,col+2,col_abs=True)}="",0,{Utils.rowcol_to_cell(row,col+2,col_abs=True)})'
            f += ")"

            self.workbook.write(
                row,
                col + 4,
                Formula(f),
                num_format_str=num_format,
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
            f += f'"корректировка начисленного резерва", ' + \
                ('"резерв по основному долгу"' if sub ==
                 "main" else '"резерв по процентам"')
            f += ")"
            self.workbook.write(
                row,
                col + 9,
                Formula(f),
            )

            row += 1

        row, col = 0, 0
        if report.is_archi:
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
        row += 1
        num_format = "#,##0.00"
        for client in report.clients.values():
            for order in client.orders:
                __write("main")
            for order in client.orders:
                __write("proc")


# , count_days:int, pdn:float, rest: float, count_delay ):
def reservePersent(order):    
    return order.percent
    


        # f = (
        #     ""
        #     + f'IF(AND({Utils.rowcol_to_cell(row,col,col_abs=True)}="",{Utils.rowcol_to_cell(row,col1,col_abs=True)}=""),"-1",'
        #     + f'IF({Utils.rowcol_to_cell(row,col1,col_abs=True)}="","-2",'
        #     + f'IF(AND({Utils.rowcol_to_cell(row,col,col_abs=True)}="",'
        #     + f'{Utils.rowcol_to_cell(row,get_col("calc_debet_end_main"),col_abs=True)}>=10000,'
        #     + f'{Utils.rowcol_to_cell(row,get_col("pdn"),col_abs=True)}>=0.5),0,'
        #     + f'IF(AND({Utils.rowcol_to_cell(row,col,col_abs=True)}="",'
        #     + f'OR({Utils.rowcol_to_cell(row,get_col("calc_debet_end_main"),col_abs=True)}<10000,'
        #     + f'{Utils.rowcol_to_cell(row,get_col("pdn"),col_abs=True)}<0.5)),"",'
        #     + f"IF(AND({Utils.rowcol_to_cell(row,col,col_abs=True)}<=7,{Utils.rowcol_to_cell(row,col-1,col_abs=True)}>=0),0,"
        #     + f"IF({Utils.rowcol_to_cell(row,col,col_abs=True)}<=30,3/100,"
        #     + f"IF({Utils.rowcol_to_cell(row,col,col_abs=True)}<=60,10/100,"
        #     + f"IF({Utils.rowcol_to_cell(row,col,col_abs=True)}<=90,20/100,"
        #     + f"IF({Utils.rowcol_to_cell(row,col,col_abs=True)}<=120,40/100,"
        #     + f"IF({Utils.rowcol_to_cell(row,col,col_abs=True)}<=180,50/100,"
        #     + f"IF({Utils.rowcol_to_cell(row,col,col_abs=True)}<=270,65/100,"
        #     + f"IF({Utils.rowcol_to_cell(row,col,col_abs=True)}<=360,80/100,"
        #     + f"IF({Utils.rowcol_to_cell(row,col,col_abs=True)}>360,99/100,"
        #     + f'IF({Utils.rowcol_to_cell(row,get_col("count_days"),col_abs=True)}>31,0,'
        #     + f'"-3"))))))))))))))'

        # )


# from xlwt import Utils
# print Utils.rowcol_pair_to_cellrange(2,2,12,2)
# print Utils.rowcol_to_cell(13,2)
#  ws.write(i, 2, xlwt.Formula("$A$%d+$B$%d" % (i+1, i+1)))
