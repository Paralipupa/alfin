from xlwt import Utils, Formula

from module.data import *


# %% Средневзвешенная
def write_result_weighted_average(self, wa: dict):
    if len(wa) == 0:
        return

    self.workbook.addSheet("Ср.взвешенная")

    names = [
        {"name": "stavka", "title": "Ставка"},
        {"name": "period", "title": "Срок"},
        {"name": "koef", "title": "Коэфф."},
    ]
    # pattern_style = "pattern: pattern solid, fore_colour green; font: color yellow; height 260;"
    pattern_style = "font: height 160;"
    pattern_style_wa = "pattern: pattern solid, fore_colour yellow; font: color black; "
    num_format_0 = "#,##0"
    num_format_2 = "#,##0.00"
    num_format_3 = "#,##0.000"
    index = 0
    items = list()
    for key, value in wa.items():
        if key.find("_30") != -1:
            items.append({"key": key, "value": value})
    len_30 = len(items)
    for key, value in wa.items():
        if key.find("_31") != -1:
            items.append({"key": key, "value": value})
    for item in items:
        index += 1
        row = 0
        col = (index - 1) * 5 + 1
        if isinstance(item["value"], dict):
            self.workbook.write(row, col - 1, item["key"].split("_")[0])
            for name in names:
                row += 1
                self.workbook.write(
                    row,
                    col - 1,
                    name["title"],
                    style_string=pattern_style,
                )
                self.workbook.write(
                    row,
                    col,
                    item["value"][name["name"]],
                )
            self.workbook.write(
                row + 1,
                col - 1,
                "Сумма",
                style_string=pattern_style,
            )
            self.workbook.write(
                row + 1,
                col,
                Formula(
                    f"SUM({Utils.rowcol_pair_to_cellrange(row+5,col+2,row+5+len(item['value']['value'])-1,col+2)})"
                ),
                num_format_str=num_format_2,
            )
            self.workbook.write(
                row + 2,
                col - 1,
                "Сумма(ср.вз.)",
                style_string=pattern_style,
            )
            self.workbook.write(
                row + 2,
                col,
                Formula(
                    f"SUM({Utils.rowcol_pair_to_cellrange(row+5,col+3,row+5+len(item['value']['value'])-1,col+3)})"
                ),
                num_format_str=num_format_2,
            )
            self.workbook.write(
                row + 3,
                col - 1,
                "Кол-во",
                style_string=pattern_style,
            )
            self.workbook.write(
                row + 3,
                col,
                Formula(
                    f"SUM({Utils.rowcol_pair_to_cellrange(row+5,col,row+5+len(item['value']['value'])-1,col)})"
                ),
            )
            sorted_values = sorted(
                item["value"]["value"].items(), key=lambda x: float(x[0])
            )
            row += 3
            row_start = row
            row += 1
            for val in sorted_values:
                row += 1
                self.workbook.write(
                    row,
                    col,
                    float(val[1]),
                    style_string=pattern_style,
                )
                self.workbook.write(
                    row,
                    col + 1,
                    float(val[0]),
                    style_string=pattern_style,
                )
                self.workbook.write(
                    row,
                    col + 2,
                    Formula(
                        f"{Utils.rowcol_to_cell(row,col)}*{Utils.rowcol_to_cell(row,col+1)}"
                    ),
                    style_string=pattern_style,
                )
                self.workbook.write(
                    row,
                    col + 3,
                    Formula(
                        f"{Utils.rowcol_to_cell(row,col+2)}*{Utils.rowcol_to_cell(3,col)}"
                    ),
                    style_string=pattern_style,
                )
            self.workbook.write(
                row + 1,
                col + 2,
                Formula(
                    f"SUM({Utils.rowcol_pair_to_cellrange(row_start+2,col+2,row,col+2)})"
                ),
                num_format_str=num_format_2,
                style_string=pattern_style,
            )
            self.workbook.write(
                row + 1,
                col + 3,
                Formula(
                    f"SUM({Utils.rowcol_pair_to_cellrange(row_start+2,col+3,row,col+3)})"
                ),
                num_format_str=num_format_2,
                style_string=pattern_style,
            )
            row += 2
            for order in item["value"]["parent"]:
                self.workbook.write(
                    row,
                    col - 1,
                    Formula(order["order"].link["name_address"])
                    if order["order"].link.get("name_address")
                    else order["order"].client.name,
                    style_string=pattern_style,
                )
                self.workbook.write(
                    row,
                    col,
                    Formula(order["order"].link["number_address"])
                    if order["order"].link.get("number_address")
                    else order["order"].number,
                    style_string=pattern_style,
                )
                self.workbook.write(
                    row,
                    col + 1,
                    Formula(order["order"].link["summa_address"])
                    if order["order"].link.get("summa_address")
                    else order["order"].summa,
                    num_format_str=num_format_2,
                    style_string=pattern_style,
                )
                self.workbook.write(
                    row,
                    col + 2,
                    order["period"],
                    num_format_str=num_format_0,
                    style_string=pattern_style,
                )
                row += 1
    self.workbook.write(
        0,
        2,
        "Общая сумма",
        style_string=pattern_style,
    )
    self.workbook.write(
        0,
        3,
        Formula(f"SUM({Utils.rowcol_pair_to_cellrange(4,0,4,len(wa)*5)})"),
        num_format_str=num_format_2,
    )
    self.workbook.write(
        1,
        2,
        "Общая сумма(ср.вз.)",
        style_string=pattern_style,
    )
    self.workbook.write(
        1,
        3,
        Formula(f"SUM({Utils.rowcol_pair_to_cellrange(5,0,5,len(wa)*5)})"),
        num_format_str=num_format_2,
    )
    self.workbook.write(
        2,
        2,
        "Сред.взвеш.",
        style_string=pattern_style,
    )
    self.workbook.write(
        2,
        3,
        Formula(
            f"SUM({Utils.rowcol_pair_to_cellrange(5,0,5,len(wa)*5)})/SUM({Utils.rowcol_pair_to_cellrange(4,0,4,len(wa)*5)})"
        ),
        # Formula(
        #     f"SUM({Utils.rowcol_pair_to_cellrange(6,0,6,len(wa)*5)})/COUNT({Utils.rowcol_pair_to_cellrange(6,0,6,len(wa)*5)})"
        # ),
        style_string=pattern_style_wa,
        num_format_str=num_format_3,
    )
