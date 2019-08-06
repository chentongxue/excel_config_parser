# coding=utf-8
import os

from tools.core.base_excel import BaseExcel
from tools.sheets.csharp_sheet import CSharpSheet


class ExcelToBinaryStream(BaseExcel):
    def __init__(self, path, export_bytes_dir, export_csharp_dir):
        super(ExcelToBinaryStream, self).__init__(path)
        self.export_bytes_dir = export_bytes_dir
        self.export_csharp_dir = export_csharp_dir

    def export_data(self):
        excel_filename = os.path.basename(self.path).replace('.xlsx', '')
        try:
            excel_filename = excel_filename.decode('utf-8')
        except UnicodeDecodeError:
            excel_filename = excel_filename.decode('gb18030')
        for sheet_parser in self.parse_sheet(CSharpSheet):
            sheet_parser.save_bytes(self.export_bytes_dir)
            sheet_parser.save_csharp(self.export_csharp_dir, excel_filename)


if __name__ == '__main__':
    test = ExcelToBinaryStream('/Users/mac/work/godus/trunk/develop/excel_to_config/excel/成就表.xlsx', 'test', 'csharp')
    test.export_data()
