# coding=utf-8
import os
import json
from tools.core.base_excel import BaseExcel


class ExcelToJson(BaseExcel):
    def __init__(self, path, export_dir_path):
        super(ExcelToJson, self).__init__(path)
        self.export_dir_path = export_dir_path

    def save_file(self, filename, data):
        filename = '%s.json' % filename
        path = os.path.join(self.export_dir_path, filename)
        text = json.dumps(data, ensure_ascii=False, indent=4)
        with open(path, 'w') as f:
            f.write(text.encode('utf-8'))

    def export_data(self):
        for sheet_parser in self.parse_sheet():
            data_list = []
            grid_list = sheet_parser.grid_list
            for row in grid_list:
                data = dict()
                for grid in row:
                    data[grid.key] = grid.value
                data_list.append(data)
            self.save_file(sheet_parser.name, data_list)


if __name__ == '__main__':
    test = ExcelToJson(u'/Users/mac/work/godus/trunk/develop/excel_to_config/excel/成就表.xlsx', '/Users/mac/work/godus/trunk/develop/config_tools/test')
    test.export_data()
