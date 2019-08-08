# coding=utf-8
import os

from tools import const
from tools.core.base_sheet import BaseSheet
from tools.core.stream_tool import StreamTool
from tools.core.csharp_code import CSharpCode


class CSharpSheet(BaseSheet):
    def __init__(self, sheet):
        super(CSharpSheet, self).__init__(sheet)
        self.stream_tool = StreamTool()

    def make_bytes(self):
        self.stream_tool.write_int(len(self.grid_list))
        for row in self.grid_list:
            for col_index, grid in enumerate(row):
                if col_index == 0:
                    self.stream_tool.write_int(grid.value)

                if grid.data_type == const.BASE_DATA_TYPES['int']:
                    self.stream_tool.write_int(grid.value)
                elif grid.data_type == const.BASE_DATA_TYPES['float']:
                    self.stream_tool.write_float(grid.value)
                elif grid.data_type == const.BASE_DATA_TYPES['str']:
                    self.stream_tool.write_str(grid.value)
                elif grid.data_type == const.BASE_DATA_TYPES['arrInt']:
                    self.stream_tool.write_arrInt(grid.value)
                elif grid.data_type == const.BASE_DATA_TYPES['arrFloat']:
                    self.stream_tool.write_arrFloat(grid.value)
                elif grid.data_type == const.BASE_DATA_TYPES['arrStr']:
                    self.stream_tool.write_arrStr(grid.value)

        return self.stream_tool.get_bytes()

    def save_bytes(self, folder_path):
        binary_stream = self.make_bytes()
        filename = '%s.bytes' % self.name
        path = os.path.join(folder_path, filename)
        with open(path, 'wb') as f:
            f.write(binary_stream)

    def save_csharp(self, folder_path, excel_filename):
        csharp_code = CSharpCode(self.name, self.data_types, self.keys, self.descriptions)
        csharp_file_content = csharp_code.get_code(excel_filename)

        filename = '%s.cs' % self.name
        path = os.path.join(folder_path, filename)
        with open(path, 'w') as f:
            f.write(csharp_file_content.encode('utf-8'))
