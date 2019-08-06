# coding=utf-8
from openpyxl import load_workbook
from tools.core.base_sheet import BaseSheet


class BaseExcel(object):
    def __init__(self, path):
        self.path = path
        self.excel = load_workbook(path, data_only=True)

    def parse_sheet(self, sheet_class=None):
        """
        :param sheet_class: 可动态传入 sheet 处理类，否则使用 BaseSheet
        :return:
        """
        _class = sheet_class or BaseSheet
        for sheet in self.excel.worksheets:
            # 只处理 Cfg 结尾的sheet
            if not sheet.title.endswith('Cfg'):
                continue
            sheet_parser = _class(sheet)
            sheet_parser.handle()
            yield sheet_parser


if __name__ == '__main__':
    test = BaseExcel(u'/Users/mac/work/godus/trunk/develop/excel_to_config/excel/成就表.xlsx')
    for parser in test.parse_sheet():
        pass
