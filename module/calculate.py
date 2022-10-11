import datetime, zipfile, pathlib
from module.oborot58 import Oborot58
from module.oborot76 import Oborot76
from module.pdn import Pdn
from module.irkom import Irkom

class Calc:
    def __init__(self, file_58: str = '', file_76: str = '', file_pdn: str = '', file_irkom: str = ''):
        self.report58 = Oborot58(file_58, '58')
        self.report76 = Oborot76(file_76, '58')
        self.reportPDN = Pdn(file_pdn, '58')
        self.reportIRKOM = Irkom(file_irkom, '58')

    def read(self):
        self.report58.get_parser()
        self.report76.get_parser()
        self.reportIRKOM.get_parser()
        self.reportPDN.get_parser()
        self.report58.union_all(self.report76, self.reportPDN, self.reportIRKOM)
        self.report58.set_weighted_average()
        self.report58.set_kategoria()

    def write(self):
        return self.report58.write_to_excel()

    def __make_archive(self, file_output: str) -> str:
        filename_arch = f'{file_output}.zip'
        arch_zip = zipfile.ZipFile(filename_arch, 'w')
        arch_zip.write(file_output, compress_type=zipfile.ZIP_DEFLATED)
        arch_zip.close()
        return filename_arch

