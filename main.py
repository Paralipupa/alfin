import pathlib
from module.calculate_async import Calc

if __name__ == '__main__':
    files=[]

    files.append(pathlib.Path('input', 'barguzin', 'ПДН 3 квартал 2022.xlsx'))
    files.append(pathlib.Path('input', 'barguzin', '58,03.xlsx'))
    files.append(pathlib.Path('input', 'barguzin', '58,03рез1.xlsx'))

    # file = pathlib.Path('input', 'report', '58,03 пдн.xlsx')
    # files.append(pathlib.Path('input', 'report', '76,06 пдн.xlsx'))
    # files.append(pathlib.Path('input', 'report', 'Отчет по ПДН.xls'))
    # files.append(pathlib.Path(
    #     'input', 'report', 'report_loan_issuance_for_prelovskaya_o (1).xls')) 
    calc = Calc(files)
    calc.read()
    calc.report_kategoria()
    calc.report_weighted_average()
    calc.write()


