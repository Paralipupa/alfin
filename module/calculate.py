from module.report import Report

class Calc:
    def __init__(self, file_58: str = '', file_76: str = '', file_pdn: str = '', file_irkom: str = ''):
        self.report58 = Report(file_58)
        self.report76 = Report(file_76, 'proc')
        self.reportPDN = Report(file_pdn)
        self.reportIRKOM = Report(file_irkom)

    def read(self):
        self.report58.get_parser()
        self.report76.get_parser()
        self.reportPDN.get_parser()
        self.reportIRKOM.get_parser()
        self.report58.union_all(self.report76, self.reportPDN, self.reportIRKOM)        
        self.report58.set_weighted_average()
        self.report58.set_kategoria()

    def write(self):
        return self.report58.write_to_excel()


