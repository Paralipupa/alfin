from re import I
from module.oborot58 import Oborot58
from module.pdn import Pdn
from module.irkom import Irkom

file_58 = 'input/report/оборотка 58,03 для пдн.xlsx'
file_pdn = 'input/report/Отчет по ПДН.xls'
file_irk = 'input/report/report_loan_issuance_for_prelovskaya_o (1).xls'
if __name__ == '__main__':
    report58 = Oborot58(file_58)
    report58.get_parser()
    reportPdn = Pdn(file_pdn)
    reportPdn.get_parser()
    reportIrk = Irkom(file_irk)
    reportIrk.get_parser()
    report58.union_all(reportPdn, reportIrk)
    report58.write('docs','docs')
    report58.weighted_average()
    report58.write('wa','result')
    report58.write_full_csv()
