# coding=utf-8
import os
import struct
from tools.template import CSHARP_TEMPLATE

from tools.core.base_sheet import BaseSheet
from tools import const


class CSharpSheet(BaseSheet):
    def __init__(self, sheet):
        super(CSharpSheet, self).__init__(sheet)

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

    def get_struct_format(self, row):
        struct_format = 'i'
        for grid in row:
            if grid.data_type == const.BASE_DATA_TYPES['int']:
                struct_format += 'i'
            elif grid.data_type == const.BASE_DATA_TYPES['float']:
                struct_format += 'f'
            elif grid.data_type == const.BASE_DATA_TYPES['str']:
                struct_format += 'i'
                str_len = max(1, len(grid.value.encode('utf-8')))
                struct_format += '%ss' % str_len

            # 如果是arr类型且长度为空时，直接写入长度0。非空时，写入n个对应类型字节
            # eg: [1,2,3]写作 iiii , []写作i
            elif grid.data_type == const.BASE_DATA_TYPES['arrInt']:
                array_len = len(grid.value)
                struct_format += 'i'
                if array_len > 0:
                    struct_format += '%si' % array_len
            elif grid.data_type == const.BASE_DATA_TYPES['arrFloat']:
                array_len = len(grid.value)
                struct_format += 'i'
                if array_len > 0:
                    struct_format += '%sf' % array_len
            elif grid.data_type == const.BASE_DATA_TYPES['arrStr']:
                struct_format += 'i'
                for item in grid.value:
                    item = item.encode('utf-8')
                    struct_format += 'i'
                    struct_format += '%ss' % len(item)
        return struct_format

    def make_stream_list(self, row):
        stream_list = []
        for col_index, grid in enumerate(row):
            if col_index == 0:
                stream_list.append(grid.value)
            if grid.data_type == const.BASE_DATA_TYPES['int']:
                stream_list.append(grid.value)
            elif grid.data_type == const.BASE_DATA_TYPES['float']:
                stream_list.append(grid.value)
            elif grid.data_type == const.BASE_DATA_TYPES['str']:
                # 字符串类型，先写入一个字符串长度的int，然后写入该字符串，如果字符串长度为0，写入空字符串''
                s = grid.value.encode('utf-8')
                str_len = max(len(s), 1)
                stream_list.append(str_len)
                stream_list.append(grid.value.encode('utf-8'))
            elif grid.data_type == const.BASE_DATA_TYPES['arrInt']:
                stream_list.append(len(grid.value))
                for item in grid.value:
                    stream_list.append(item)
            elif grid.data_type == const.BASE_DATA_TYPES['arrFloat']:
                stream_list.append(len(grid.value))
                for item in grid.value:
                    stream_list.append(item)
            elif grid.data_type == const.BASE_DATA_TYPES['arrStr']:
                # arrStr 类型，先写入数组长度，然后写入第一个元素的长度，最后写入第一个元素的值，以此类推
                stream_list.append(len(grid.value))
                for item in grid.value:
                    item = item.encode('utf-8')
                    stream_list.append(len(item))  # 写入元素长度
                    stream_list.append(item)       # 写入值
        return stream_list

    def save_bytes(self, folder_path):
        sheet_struct_format = '<i'
        sheet_stream_list = [len(self.grid_list)]

        for row in self.grid_list:
            sheet_struct_format += self.get_struct_format(row)
            sheet_stream_list.extend(self.make_stream_list(row))
        sheet_binary_stream = struct.pack(sheet_struct_format, *sheet_stream_list)

        filename = '%s.bytes' % self.name
        path = os.path.join(folder_path, filename)
        with open(path, 'w') as f:
            f.write(sheet_binary_stream)

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
