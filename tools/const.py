# coding=utf-8
import struct


BASE_DATA_TYPES = {
    'int': 0,
    'float': 1,
    'str': 2,
    'arrInt': 3,
    'arrFloat': 4,
    'arrStr': 5,
}


DATA_TYPE_STRUCT_LEN = {
    BASE_DATA_TYPES['int']: struct.calcsize('i'),
    BASE_DATA_TYPES['float']: struct.calcsize('f'),
    BASE_DATA_TYPES['str']: struct.calcsize('s')
}


DESCRIPTION_ROW_NUM = 0  # 描述行号
COMMENT_ROW_NUM = 1      # 注释行号
KEY_ROW_NUM = 2          # 键行号
SCOPE_ROW_NUM = 3        # 作用范围行号
DATA_TYPE_ROW_NUM = 4    # 数据类型行号
DATA_START_ROW_NUM = 5   # 数据开始行号

SCOPE_SERVER = 'S'  # 服务器
SCOPE_CLIENT = 'C'  # 客户端
