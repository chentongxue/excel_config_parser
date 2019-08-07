# coding=utf-8
from tools import const
from tools.template import CSHARP_TEMPLATE


define_indent = '\t\t\t'
decode_indent = '\t\t\t\t'


class CSharpCode(object):
    def __init__(self, name, data_types, keys, descriptions):
        self.name = name
        self.key_defines = []
        self.decode_codes = []
        self.data_types = data_types
        self.keys = keys
        self.descriptions = descriptions

    def write_int_define(self, key):
        self.key_defines.append(define_indent + 'public int %s;' % key)

    def write_float_define(self, key):
        self.key_defines.append(define_indent + 'public float %s;' % key)

    def write_str_define(self, key):
        self.key_defines.append(define_indent + 'public string %s;' % key)

    def write_arr_int_define(self, key):
        self.key_defines.append(define_indent + 'public List<int> %s = new List<int>();' % key)

    def write_arr_float_define(self, key):
        self.key_defines.append(define_indent + 'public List<float> %s = new List<float>();' % key)

    def write_arr_str_define(self, key):
        self.key_defines.append(define_indent + 'public List<string> %s = new List<string>();' % key)

    def write_int_decode(self, key):
        self.decode_codes.append(decode_indent + '%s = In.ReadInt32();' % key)

    def write_float_decode(self, key):
        self.decode_codes.append(decode_indent + '%s = In.ReadSingle();' % key)

    def write_str_decode(self, key):
        self.decode_codes.append(decode_indent + '%s = st.ReadAsciiString();' % key)

    def write_arr_int_decode(self, key):
        self.decode_codes.append(decode_indent + '%s.Clear();' % key)
        s = decode_indent + """var n_arrInt = In.ReadInt32();
                for (int i = 0, n = n_arrInt; i < n; ++i)
                {
                    var a = In.ReadInt32();
                    %s.Add(a);
                }""" % key
        self.decode_codes.append(s)

    def write_arr_float_decode(self, key):
        self.decode_codes.append(decode_indent + '%s.Clear();' % key)
        s = decode_indent + """var n_arrFloat = In.ReadInt32();
                for (int i = 0, n = n_arrFloat; i < n; ++i)
                {
                    %s.Add(In.ReadSingle());
                }""" % key
        self.decode_codes.append(s)

    def write_arr_str_decode(self, key):
        self.decode_codes.append(decode_indent + '%s.Clear();' % key)
        s = decode_indent + """var n_arrStr = In.ReadInt32();
                for (int i = 0, n = n_arrStr; i < n; ++i)
                {
                    %s.Add(st.ReadAsciiString());
                }""" % key
        self.decode_codes.append(s)

    def make_code_array(self):
        for index, data_type in enumerate(self.data_types):
            key = self.keys[index]
            self.key_defines.append(define_indent + '/// <summary> %s </summary>' % self.descriptions[index])
            if data_type == const.BASE_DATA_TYPES['int']:
                self.write_int_define(key)
                self.write_int_decode(key)
            elif data_type == const.BASE_DATA_TYPES['float']:
                self.write_float_define(key)
                self.write_float_decode(key)
            elif data_type == const.BASE_DATA_TYPES['str']:
                self.write_str_define(key)
                self.write_str_decode(key)
            elif data_type == const.BASE_DATA_TYPES['arrInt']:
                self.write_arr_int_define(key)
                self.write_arr_int_decode(key)
            elif data_type == const.BASE_DATA_TYPES['arrFloat']:
                self.write_arr_float_define(key)
                self.write_arr_float_decode(key)
            elif data_type == const.BASE_DATA_TYPES['arrStr']:
                self.write_arr_str_define(key)
                self.write_arr_str_decode(key)

    def get_code(self, excel_filename):
        self.make_code_array()
        decode_codes = '\n'.join(self.decode_codes)
        key_defines = '\n'.join(self.key_defines)
        code = CSHARP_TEMPLATE.replace('##SHEET_NAME##', self.name)
        code = code.replace('##DECODE##', decode_codes)
        code = code.replace('##KEY_DEFINE##', key_defines)
        code = code.replace('##EXCEL_FILENAME##', excel_filename)
        return code

