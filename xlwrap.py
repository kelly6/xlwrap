#!/usr/bin/python
# -*- coding: UTF-8 -*-
import traceback
import xlwt
import xlrd

class xl_loader():
    def __init__(self, fpath):
        self.fpath = fpath

    def load(self):
        sheets_dic = {}
        wb = xlrd.open_workbook(self.fpath)
        sheets = wb.sheets()
        for sheet in sheets:
            matrix = []
            nrows = sheet.nrows
            for i in range(nrows):
                l = sheet.row_values(i)
                l = [i.encode("utf8") if isinstance(i, unicode) else i for i in l]
                matrix.append(l)
            name = sheet.name.encode("utf8") if isinstance(sheet.name, unicode) else sheet.name
            sheets_dic[name] = matrix
        return sheets_dic

class xl_writer():
    def __init__(self, sheets_dic):
        self.sheets_dic = sheets_dic
    
    def save(self, fpath):
        wb = xlwt.Workbook()
        for k in self.sheets_dic:
            sheet = wb.add_sheet(k if isinstance(k, unicode) else k.decode("utf8"))
            for row_idx, row in enumerate(self.sheets_dic[k]):
                for col_idx, cell in enumerate(row):
                    sheet.write(row_idx, col_idx, cell.decode("utf8") if isinstance(cell, str) else cell)
        wb.save(fpath)

if __name__ == "__main__":
    loader = xl_loader("./test.xls")

    d = loader.load()
    print d.keys()

    writer = xl_writer(d)
    writer.save("/tmp/test.bak.xls")
