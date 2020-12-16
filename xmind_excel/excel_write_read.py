# -*- coding: utf-8 -*-
"""
Excel 文件流处理
"""
from  openpyxl import  Workbook
class ExcelWriteRead(object):
    def __init__(self, path, file_name, sheet_name):
        # self.path = path
        self.file_name = file_name
        self.sheet_name = sheet_name

        self.wb = Workbook()
        self.ws = self.wb.active

    def excel_write(self):
        # self.ws.title = "ame"
        ws1 = self.wb.create_sheet("rwerwe")
        ws2 = self.wb.create_sheet("02")

        self.wb.save(self.file_name)

if __name__ == '__main__':
    n = ExcelWriteRead('/Users/wuting/wt测试.xls', "/Users/wuting/wt测试002.xlsx","wt")

    n.excel_write()