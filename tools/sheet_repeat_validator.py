# coding=utf-8
from collections import Counter


class SheetRepeatValidator(object):
    def __init__(self):
        self.excel_sheet_names = dict()
        self.sheet_names = []
        self.error_messages = []

    def add_name(self, sheet_name, excel_filename):
        self.excel_sheet_names.setdefault(excel_filename, [])
        self.excel_sheet_names[excel_filename].append(sheet_name)
        self.sheet_names.append(sheet_name)

    def check_repeat(self):
        for name, count in Counter(self.sheet_names).iteritems():
            if count <= 1:
                continue
            for excel_filename, sheet_names in self.excel_sheet_names.iteritems():
                for sheet_name in sheet_names:
                    if sheet_name == name:
                        self.error_messages.append(u'%s 存在重名表 %s' % (excel_filename, sheet_name))


if __name__ == '__main__':
    validator = SheetRepeatValidator()
    validator.add_name('testCfg', 'a.xlsx')
    validator.add_name('testCfg', 'b.xlsx')
    validator.add_name('test1Cfg', 'a.xlsx')

    validator.check_repeat()
