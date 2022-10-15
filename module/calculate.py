from module.report import Report

class Calc:
    def __init__(self, file_name: str, *args):
        self.main = Report(file_name)
        self.items=[]
        for name in args:
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


