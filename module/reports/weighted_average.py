from xlwt import Utils, Formula

from module.data import *

# %% Средневзвешенная
def write_result_weighted_average(self, result: dict):
    if len(result) == 0:
        return
    
    self.workbook.addSheet("Ср.взвешенная")
    
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
                    col - 1,
                    Formula(order.link["name_address"])
                    if order.link.get("name_address")
                    else order.client.name,
                )
                self.workbook.write(
                    row,
                    col,
                    Formula(order.link["number_address"])
                    if order.link.get("number_address")
                    else order.number,
                )
                self.workbook.write(
                    row,
                    col + 1,
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
