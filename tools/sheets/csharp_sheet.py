# coding=utf-8
import os
from tools.template import CSHARP_TEMPLATE

from tools.core.base_sheet import BaseSheet
from tools import const
from tools.core.stream_tool import StreamTool


class CSharpSheet(BaseSheet):
    def __init__(self, sheet):
        super(CSharpSheet, self).__init__(sheet)
        self.stream_tool = StreamTool()

    def make_csharp_key_define(self):
        codes = []
        indent = '\t\t\t'
        for index, data_type in enumerate(self.data_types):
            key = self.keys[index]
            codes.append(indent + '/// <summary> %s </summary>' % self.descriptions[index])
            if data_type == const.BASE_DATA_TYPES['int']:
                codes.append(indent + 'public int %s;' % key)
            elif data_type == const.BASE_DATA_TYPES['float']:
                codes.append(indent + 'public float %s;' % key)
            elif data_type == const.BASE_DATA_TYPES['str']:
                codes.append(indent + 'public string %s;' % key)
            elif data_type == const.BASE_DATA_TYPES['arrInt']:
                codes.append(indent + 'public List<int> %s = new List<int>();' % key)
            elif data_type == const.BASE_DATA_TYPES['arrFloat']:
                codes.append(indent + 'public List<float> %s = new List<float>();' % key)
            elif data_type == const.BASE_DATA_TYPES['arrStr']:
                codes.append(indent + 'public List<string> %s = new List<string>();' % key)
        return '\n'.join(codes)

    def make_csharp_decode(self):
        codes = []
        indent = '\t\t\t\t'
        for index, data_type in enumerate(self.data_types):
            key = self.keys[index]
            if data_type == const.BASE_DATA_TYPES['int']:
                codes.append(indent + '%s = In.ReadInt32();' % key)
            elif data_type == const.BASE_DATA_TYPES['float']:
                codes.append(indent + '%s = In.ReadSingle();' % key)
            elif data_type == const.BASE_DATA_TYPES['str']:
                codes.append(indent + '%s = st.ReadAsciiString();' % key)
            elif data_type == const.BASE_DATA_TYPES['arrInt']:
                codes.append(indent + '%s.Clear();' % key)
                s = indent + """var n_arrInt = In.ReadInt32();
                for (int i = 0, n = n_arrInt; i < n; ++i)
                {
                    var a = In.ReadInt32();
                    %s.Add(a);
                }""" % key
                codes.append(s)
            elif data_type == const.BASE_DATA_TYPES['arrFloat']:
                codes.append(indent + '%s.Clear();' % key)
                s = indent + """var n_arrFloat = In.ReadInt32();
                for (int i = 0, n = n_arrFloat; i < n; ++i)
                {
                    %s.Add(In.ReadSingle());
                }""" % key
                codes.append(s)
            elif data_type == const.BASE_DATA_TYPES['arrStr']:
                codes.append(indent + '%s.Clear();' % key)
                s = indent + """var n_arrStr = In.ReadInt32();
                for (int i = 0, n = n_arrStr; i < n; ++i)
                {
                    %s.Add(st.ReadAsciiString());
                }""" % key
                codes.append(s)
        return '\n'.join(codes)

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
        with open(path, 'w') as f:
            f.write(binary_stream)

    def save_csharp(self, folder_path, excel_filename):
        csharp_decode = self.make_csharp_decode()
        csharp_key_defines = self.make_csharp_key_define()
        csharp_file_content = CSHARP_TEMPLATE.replace('##SHEET_NAME##', self.name) \
            .replace('##DECODE##', csharp_decode) \
            .replace('##KEY_DEFINE##', csharp_key_defines) \
            .replace('##EXCEL_FILENAME##', excel_filename)

        filename = '%s.cs' % self.name
        path = os.path.join(folder_path, filename)
        with open(path, 'w') as f:
            f.write(csharp_file_content.encode('utf-8'))
