from module.report import Report

class Calc:
    def __init__(self, files: list):
        self.main_s = Report(files[0])
        self.main_p = Report(files[1])
        self.items=[]
        for name in files[2:]:
            self.items.append(Report(name))

    def read(self):
        self.main_s.get_parser()
        self.main_p.get_parser()
        for rep in self.items:
            rep.get_parser()
        self.main_s.union_all(*self.items)
        self.main_p.union_all(*self.items)
    
    def report_weighted_average(self):
        self.main_s.set_weighted_average()
        self.main_p.result = self.main_s.result
    
    def report_rezerves(self):
        self.main_p.set_reserves()
        self.main_s.kategoria = self.main_p.kategoria

    def write(self):
        return self.main_p.write_to_excel()


