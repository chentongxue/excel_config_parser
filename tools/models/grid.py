# coding=utf-8
import string
import json
from tools import const
from tools.models import validator


class Grid(object):
    """
    对Excel单元格的抽象
    """
    def __init__(self, key, value, description, comment, scope, data_type, row_num, col_num):
        """
        所有 property 根据 传入的 data_type 处理成 python 类型

        :param key: 原始key
        :param value: 原始value
        :param description:  原始描述
        :param comment: 原始注释
        :param scope: 原始作用域
        :param data_type: 原始数据类型
        :param row_num: excel中的行号
        :param col_num: excel中的列号
        """
        self.raw_key = key
        self.raw_value = value
        self.raw_description = description
        self.raw_comment = comment
        self.raw_scope = scope
        self.raw_data_type = data_type
        self.row_num = row_num
        self.col_num = col_num

    def __repr__(self):
        return ('<Grid %s: %s [%s]>' % (self.position, self.raw_value, self.raw_data_type)).encode('utf-8')

    @property
    def key(self):
        return self.raw_key

    @property
    def value(self):
        """
        预处理原始值
        对空值返回默认值
        :return:
        """
        if self.raw_value is None:
            if self.data_type == const.BASE_DATA_TYPES['int']:
                return 0
            elif self.data_type == const.BASE_DATA_TYPES['float']:
                return 0.0
            elif self.data_type == const.BASE_DATA_TYPES['str']:
                return ''
            elif self.data_type in [const.BASE_DATA_TYPES['arrInt'], const.BASE_DATA_TYPES['arrFloat'], const.BASE_DATA_TYPES['arrStr']]:
                return []
        else:
            if self.data_type in [const.BASE_DATA_TYPES['arrInt'], const.BASE_DATA_TYPES['arrFloat'], const.BASE_DATA_TYPES['arrStr']]:
                return json.loads(self.raw_value)
            elif self.data_type == const.BASE_DATA_TYPES['str']:
                try:
                    return str(self.raw_value)
                except UnicodeEncodeError:
                    return self.raw_value
        return self.raw_value

    @property
    def description(self):
        return self.raw_description

    @property
    def comment(self):
        return self.raw_comment

    @property
    def scope(self):
        return self.raw_scope

    @property
    def data_type(self):
        """
        数据类型对应的int值
        :return: int
        """
        return const.BASE_DATA_TYPES[self.raw_data_type]

    @property
    def position(self):
        """
        计算实例的excel坐标
        :return: str
        """
        letters = string.letters[:26].upper()
        arr = []
        n = self.col_num + 1
        while n:
            k = n % 26 or 26
            arr.insert(0, letters[k-1])
            n = (n - k) / 26
        grid = ''.join(arr) + str(self.row_num + 1)
        return grid

    def validate(self):
        """
        校验本单元格
        :return:
        """
        validate = validator.Validator(self.raw_data_type, self.raw_value)
        # 第1列必须为int类型，且不能为空
        if self.col_num == 0:
            validate.validate_primary_key()
            validate.validate()
            return
        if self.raw_value is None:
            return
        validate.validate()
