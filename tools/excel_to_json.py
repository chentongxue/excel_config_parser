# coding=utf-8
from tools.core.base_excel import BaseExcel
from tools.sheets.json_sheet import JsonSheet


class ExcelToJson(BaseExcel):
    def __init__(self, path, export_dir_path):
        super(ExcelToJson, self).__init__(path)
        self.export_dir_path = export_dir_path

    def export_data(self):
        for sheet_parser in self.parse_sheet(JsonSheet):
            sheet_parser.save_file(self.export_dir_path)


if __name__ == '__main__':
    test = ExcelToJson(u'/Users/mac/work/godus/trunk/develop/excel_to_config/excel/成就表.xlsx', '/Users/mac/work/godus/trunk/develop/config_tools/test')
    test.export_data()
