from re import I
from module.oborot58 import Oborot58
from module.oborot76 import Oborot76
from module.pdn import Pdn
from module.irkom import Irkom
import pathlib

file_58 = pathlib.Path('input', 'report', '58,03 пдн.xlsx')
file_76 = pathlib.Path('input', 'report', '76,06 пдн.xlsx')
file_pdn = pathlib.Path('input', 'report', 'Отчет по ПДН.xls')
file_irk = pathlib.Path(
    'input', 'report', 'report_loan_issuance_for_prelovskaya_o (1).xls')
if __name__ == '__main__':
    report58 = Oborot58(file_58, '58')
    report58.get_parser()
    report76 = Oborot76(file_76, '76')
    report76.get_parser()
    reportPdn = Pdn(file_pdn, 'PDN')
    reportPdn.get_parser()
    reportIrk = Irkom(file_irk, 'IRKOM')
    reportIrk.get_parser()
    report58.union_all(report76, reportPdn, reportIrk)
    report58.write('docs', 'docs')
    report58.set_weighted_average()
    report58.write('wa', 'result')
    report58.set_kategoria()
    report58.write('kateg', 'kategoria')
    report58.write_to_excel()
