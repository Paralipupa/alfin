from xlwt import Utils, Formula

from module.data import *


def write_clients(self, report) -> bool:
    def get_names():
        return (
            [
                {"name": "number", "title": "Номер", "type": ""},
                {"name": "summa", "title": "Сумма", "type": "float"},
            ]
            + (
                [
                    {
                        "name": "debet_end_main",
                        "title": "Дт.58",
                        "type": "float",
                    },
                    {
                        "name": "debet_end_proc",  # "calc_debet_end_proc",
                        "title": "Дт.76",
                        "type": "float",
                    },
                    {
                        "name": "calc_reserve_main",
                        "title": "Дт.59",
                        "type": "float",
                    },
                    {
                        "name": "calc_reserve_proc",
                        "title": "Дт.63",
                        "type": "float",
                    },
                    {"name": "reserve_percent", "title": "%", "type": "int"},
                    {"name": "credit_main", "title": "Кт.58", "type": "float"},
                    {
                        "name": "calc_credit_proc",
                        "title": "Кт.76",
                        "type": "float",
                    },
                    {"name": "credit_end_main", "title": "Кт.59", "type": "float"},
                    {"name": "credit_end_proc", "title": "Кт.63", "type": "float"},
                ]
                if not report.options.get("option_weighted_average")
                else []
            )
            + [
                {"name": "rate", "title": "Ставка", "type": "float"},
                {"name": "tarif", "title": "Тариф", "type": ""},
                {"name": "count_days", "title": "Срок", "type": "float"},
            ]
            + (
                [
                    {"name": "pdn", "title": "ПДН", "type": "float"},
                    {"name": "date_begin", "title": "ДатНач.", "type": "date"},
                ]
                if report.options.get("option_kategory")
                else []
            )
            + (
                [
                    # {"name": "payments_base", "title": "Archi", "type": "float"},
                    {
                        "name": "date_frozen",
                        "title": "ДатЗам.",
                        "type": "date",
                    },
                ]
                if report.options.get("option_is_archi")
                and not report.options.get("option_weighted_average")
                else []
            )
            + (
                [
                    {
                        "name": "count_days_common",
                        "title": "Дн.всего",
                        "type": "int",
                    },
                    {
                        "name": "count_days_delay",
                        "title": "Дн.пр.",
                        "type": "int",
                    },
                    # {
                    #     "name": "calculate_percent",
                    #     "title": "Проц.всего",
                    #     "type": "float",
                    # },
                    # {
                    #     "name": "count_days_period",
                    #     "title": "Дней\n(месяц)",
                    #     "type": "int",
                    # },
                    # {
                    #     "name": "summa_percent",
                    #     "title": "Проц.месяц",
                    #     "type": "float",
                    # },
                ]
                if not report.options.get("option_weighted_average")
                else []
            )
            + (
                [
                    {
                        "name": "debet_end_proc_58",
                        "title": "Дт.76(н)",
                        "type": "float",
                    },
                    {
                        "name": "calc_debet_end_proc_58",
                        "title": "",
                        "type": "float",
                    },
                    {
                        "name": "summa_reserve_main_58",
                        "title": "Дт.59(н)",
                        "type": "float",
                    },
                    {
                        "name": "calc_summa_reserve_main_58",
                        "title": "",
                        "type": "float",
                    },
                    # {"name": "summa_reserve_main_58_pdn",
                    #     "title": "резерв по основному долгу (по ставке)", "type": "float"},
                    # {
                    #     "name": "summa_reserve_proc_58",
                    #     "title": "Дт.63()",
                    #     "type": "float",
                    # },
                    # {
                    #     "name": "calc_summa_reserve_proc_58",
                    #     "title": "",
                    #     "type": "float",
                    # },
                    # {"name": "summa_reserve_proc_58_pdn",
                    #     "title": "резерв по процентам (по ставке)", "type": "float"},
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
                if report.options.get("option_is_archi")
                and not report.options.get("option_weighted_average")
                else []
            )
        )

    def write_value_attribute(value):
        nonlocal name, row, order
        try:
            if name["name"] == "date_calculate":
                if value != report.report_date:
                    self.workbook.write(row, name["col"], value, type_name=name["type"])
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
                    self.workbook.write(row, name["col"], value, type_name=name["type"])
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
            if name["name"] == "summa_percent":
                self.workbook.write(
                    row,
                    name["col"],
                    order.summa_percent_period,
                    type_name=name["type"],
                )
                # f = f"{Utils.rowcol_to_cell(row,get_col('summa'),col_abs=True)}*"
                # f += f"({Utils.rowcol_to_cell(row,get_col('rate'),col_abs=True)}/100)*"
                # f += f"{Utils.rowcol_to_cell(row,get_col('count_days_period'),col_abs=True)}"
                # self.workbook.write(
                #     row, name["col"], Formula(f), type_name=name["type"]
                # )
            elif name["name"] == "calculate_percent":
                # summa_max = order.summa * \
                #     Decimal(get_max_margin_rate(order.date))
                # f = f"MAX(MIN({summa_max},"
                # f += f"{Utils.rowcol_to_cell(row,get_col('summa'),col_abs=True)}*"
                # f += f"({Utils.rowcol_to_cell(row,get_col('rate'),col_abs=True)}/100)*"
                # f += f"{Utils.rowcol_to_cell(row,get_col('count_days_common'),col_abs=True)})-"
                # f += f"{Utils.rowcol_to_cell(row,get_col('credit_main'),col_abs=True)},0)"
                self.workbook.write(
                    row,
                    name["col"],
                    order.summa_percent_all,
                    type_name=name["type"],
                )
            elif name["name"] == "debet_end_main":
                # f = f"MAX({Utils.rowcol_to_cell(row,get_col('summa'),col_abs=True)}-"
                # f += f"{Utils.rowcol_to_cell(row,get_col('credit_main'),col_abs=True)},0)"
                self.workbook.write(
                    row, name["col"], order.debet_end_main, type_name=name["type"]
                )
                # order.debet_end_main = max(
                #     order.summa - order.credit_main, 0)
            elif name["name"] == "calc_debet_end_proc":
                # f = f"MAX({Utils.rowcol_to_cell(row,get_col('calculate_percent'),col_abs=True)}-"
                # f += f"{Utils.rowcol_to_cell(row,get_col('credit_proc'),col_abs=True)},0)"
                self.workbook.write(
                    row, name["col"], order.debet_end_proc, type_name=name["type"]
                )
                # order.debet_end_proc = max(
                #     order.calculate_percent - order.credit_proc, 0
                # )
            elif name["name"] == "reserve_percent":
                self.workbook.write(row, name["col"], order.percent)
            elif name["name"] == "calc_credit_proc":
                self.workbook.write(
                    row,
                    name["col"],
                    order.credit_proc
                    if order.summa_payment == 0
                    else order.summa_payment,
                )

            elif name["name"] == "calc_reserve_percent":
                col = get_col("count_days_delay")
                col1 = get_col("pdn")
                f = (
                    ""
                    + f'IF(AND({Utils.rowcol_to_cell(row,col,col_abs=True)}="",{Utils.rowcol_to_cell(row,col1,col_abs=True)}=""),"",'
                    + f'IF({Utils.rowcol_to_cell(row,col1,col_abs=True)}="","",'
                    + f'IF(AND({Utils.rowcol_to_cell(row,col,col_abs=True)}="",'
                    + f'{Utils.rowcol_to_cell(row,get_col("debet_end_main"),col_abs=True)}>=10000,'
                    + f'{Utils.rowcol_to_cell(row,get_col("pdn"),col_abs=True)}>=0.5),0,'
                    + f'IF(AND({Utils.rowcol_to_cell(row,col,col_abs=True)}="",'
                    + f'OR({Utils.rowcol_to_cell(row,get_col("debet_end_main"),col_abs=True)}<10000,'
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
                    f += f'{Utils.rowcol_to_cell(row,get_col("date_begin"),col_abs=True)}'
                else:
                    f = f'{Utils.rowcol_to_cell(0,0,col_abs=True,row_abs=True)}-{Utils.rowcol_to_cell(row,get_col("date_begin"),col_abs=True)}'
                self.workbook.write(
                    row, name["col"], Formula(f), type_name=name["type"]
                )
            elif name["name"] == "calc_debet_end_proc_58":
                f = f'{Utils.rowcol_to_cell(row,get_col("debet_end_proc"),col_abs=True)}-{Utils.rowcol_to_cell(row,get_col("debet_end_proc_58"),col_abs=True)}'
                self.workbook.write(
                    row, name["col"], Formula(f), type_name=name["type"]
                )
            elif name["name"] == "calc_summa_reserve_main_58":
                f = f'{Utils.rowcol_to_cell(row,get_col("calc_reserve_main"),col_abs=True)}-{Utils.rowcol_to_cell(row,get_col("summa_reserve_main_58"),col_abs=True)}'
                self.workbook.write(
                    row, name["col"], Formula(f), type_name=name["type"]
                )
            elif name["name"] == "calc_summa_reserve_proc_58":
                f = f'{Utils.rowcol_to_cell(row,get_col("calc_reserve_proc"),col_abs=True)}-{Utils.rowcol_to_cell(row,get_col("summa_reserve_proc_58"),col_abs=True)}'
                self.workbook.write(
                    row, name["col"], Formula(f), type_name=name["type"]
                )
            elif name["name"] == "count_days_delay":
                f = f'IF({Utils.rowcol_to_cell(row,get_col("count_days_common"),col_abs=True)}-'
                f += (
                    f'{Utils.rowcol_to_cell(row,get_col("count_days") if get_col("count_days") else 31,col_abs=True)}>0,'
                )
                f += f'{Utils.rowcol_to_cell(row,get_col("count_days_common"),col_abs=True)}-'
                f += f'{Utils.rowcol_to_cell(row,get_col("count_days") if get_col("count_days") else 31,col_abs=True)},'
                f += f'"")'
                self.workbook.write(
                    row, name["col"], order.count_days_delay  if get_col("count_days") else get_col("count_days_common") - 31, type_name=name["type"]
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
        # f = f'IF({Utils.rowcol_to_cell(row,get_col("reserve_percent"),col_abs=True)}="","",'
        # f += f'IF({Utils.rowcol_to_cell(row,get_col("reserve_percent"),col_abs=True)}<0,"",'
        # f += f'ROUND('
        # f += f'IF({Utils.rowcol_to_cell(row,get_col("reserve_percent"),col_abs=True)}=0,'
        # f += f'{Utils.rowcol_to_cell(row,get_col("calc_debet_end_main"),col_abs=True)}*1/10,'
        # f += f'{Utils.rowcol_to_cell(row,get_col("calc_debet_end_main"),col_abs=True)}*'
        # f += f'{Utils.rowcol_to_cell(row,get_col("reserve_percent"),col_abs=True)}'
        # f += f"),2)))"
        self.workbook.write(
            row, name["col"], order.calc_reserve_main, num_format_str=num_format
        )
        order.link[
            "calc_reserve_main" + "_address"
        ] = f"{self.workbook.sheet.name}!{Utils.rowcol_to_cell(row,len(names)+1)}"
        # order.calc_reserve_main = order.debet_end_main * \
        #     Decimal(order.percent)

    def calculate_rezerves_proc():
        # f = f'IF({Utils.rowcol_to_cell(row,get_col("reserve_percent"),col_abs=True)}="","",'
        # f += f'IF({Utils.rowcol_to_cell(row,get_col("reserve_percent"),col_abs=True)}<0,"",'
        # f += f'ROUND('
        # f += f'IF({Utils.rowcol_to_cell(row,get_col("reserve_percent"),col_abs=True)}=0,'
        # f += f'{Utils.rowcol_to_cell(row,get_col("debet_end_proc"),col_abs=True)}*1/10,'
        # f += f'{Utils.rowcol_to_cell(row,get_col("debet_end_proc"),col_abs=True)}*'
        # f += f'{Utils.rowcol_to_cell(row,get_col("reserve_percent"),col_abs=True)}'
        # f += f"),2)))"
        self.workbook.write(
            row, name["col"], order.calc_reserve_proc, num_format_str=num_format
        )
        order.link[
            "calc_reserve_proc" + "_address"
        ] = f"{self.workbook.sheet.name}!{Utils.rowcol_to_cell(row,len(names) + 2)}"
        # order.calc_reserve_proc = order.debet_end_proc * \
        #     Decimal(order.percent)

    def write_header():
        row = 1
        self.workbook.write(row, 0, "ФИО")
        for col, name in enumerate(names, 1):
            name["col"] = col
            self.workbook.write(row, col, name["title"])
        # self.workbook.write(row, len(names) + 1, "Резерв(осн.)")
        # self.workbook.write(row, len(names) + 2, "Резерв(проц.)")

    # ----------------------------------------------------------------------------------------------------------
    self.workbook.addSheet("Общий")
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
            self.workbook.write(row, len(names) + 1, client.passport_number)
            row += 1
