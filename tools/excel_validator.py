# coding=utf-8
import os

from tools.core.base_excel import BaseExcel
from tools.models import validator


class ExcelValidator(BaseExcel):
    def __init__(self, path):
        super(ExcelValidator, self).__init__(path)
        self.error_msgs = []

    def add_error_msg(self, sheet_name, tip):
        filename = os.path.basename(self.path)
        try:
            filename = filename.decode('utf-8')
        except UnicodeDecodeError:
            filename = filename.decode('gb18030')
        msg = u'%s [%s] %s' % (filename, sheet_name, tip)
        self.error_msgs.append(msg)

    def run(self):
        for sheet_parser in self.parse_sheet():
            for row in sheet_parser.grid_list:
                for grid in row:
                    try:
                        grid.validate()
                    except validator.ValidateValueError:
                        tip = u'单元格 [%s] 值填写错误' % grid.position
                        self.add_error_msg(sheet_parser.name, tip)
                    except validator.ValidateDataTypeError:
                        tip = u'表头类型设置错误, [%s]' % grid.raw_data_type
                        self.add_error_msg(sheet_parser.name, tip)
                    except validator.ValidatePrimaryKeyTypeError:
                        tip = u'主键类型设置错误, [%s], 必须为 [int]' % grid.raw_data_type
                        self.add_error_msg(sheet_parser.name, tip)

            if len(set(sheet_parser.primary_keys)) != len(sheet_parser.primary_keys):
                tip = u'存在重复主键'
                self.add_error_msg(sheet_parser.name, tip)


if __name__ == '__main__':
    test = ExcelValidator(u'/Users/mac/work/godus/trunk/develop/config_tools/excel/成就表.xlsx')
    test.run()

