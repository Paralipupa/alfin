from module.oborot58 import Oborot58
from module.pdn import Pdn

file_58 = 'input/report/оборотка 58,03 для пдн.xlsx'
file_pdn = 'input/report/Отчет по ПДН.xls'
if __name__ == '__main__':
    report58 = Oborot58(file_58)
    report58.get_parser()
    reportPdn = Pdn(file_pdn)
    reportPdn.get_parser()
