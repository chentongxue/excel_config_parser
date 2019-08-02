# coding=utf-8
from openpyxl import load_workbook
from tools.core.base_sheet import BaseSheet


class BaseExcel(object):
    def __init__(self, path):
        self.path = path
        self.excel = load_workbook(path, data_only=True)

    def parse_sheet(self):
        for sheet in self.excel.worksheets:
            sheet_parser = BaseSheet(sheet)
            sheet_parser.handle()
            yield sheet_parser


if __name__ == '__main__':
    test = BaseExcel(u'/Users/mac/work/godus/trunk/develop/excel_to_config/excel/成就表.xlsx')
    for parser in test.parse_sheet():
        pass
