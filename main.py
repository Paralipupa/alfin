import pathlib
from module.calculate import Calc

if __name__ == '__main__':
    file_58 = pathlib.Path('input', 'report', '58,03 пдн.xlsx')
    file_76 = pathlib.Path('input', 'report', '76,06 пдн.xlsx')
    file_pdn = pathlib.Path('input', 'report', 'Отчет по ПДН.xls')
    file_irkom = pathlib.Path(
        'input', 'report', 'report_loan_issuance_for_prelovskaya_o (1).xls')
    calc = Calc(file_58=file_58,file_76=file_76,file_irkom=file_irkom,file_pdn=file_pdn)
    calc.read()
    calc.write()
