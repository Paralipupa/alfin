from module.report import Report

class Calc:
    def __init__(self, files: list):
        self.main = Report(files[0])
        self.items=[]
        for name in files[1:]:
            self.items.append(Report(name))

    def read(self):
        self.main.get_parser()
        for rep in self.items:
            rep.get_parser()
        self.main.union_all(*self.items)
    
    def report_weighted_average(self):
        self.main.set_weighted_average()
    
    def report_rezerves(self):
        self.main.set_reserves()

    def write(self):
        return self.main.write_to_excel()


