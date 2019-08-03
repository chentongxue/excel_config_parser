# coding=utf-8
import json
from tools import const


class ValidateDataTypeError(Exception):
    """表头定义了不存在的类型"""


class ValidateValueError(Exception):
    """单元格值错误"""


type_int = (int, long)
type_float = (int, long, float)
type_str = (str, unicode)


class Validator(object):
    def __init__(self, data_type, value):
        self.data_type = data_type
        self.value = value

    def validate_data_type(self):
        if self.data_type not in const.BASE_DATA_TYPES:
            raise ValidateDataTypeError

    def validate_value(self):
        if self.value is None:
            return

        func = getattr(self, '_validate_value_%s' % self.data_type)
        func()

    def _validate_value_int(self):
        if type(self.value) not in type_int:
            raise ValidateValueError

    def _validate_value_float(self):
        if type(self.value) not in type_float:
            raise ValidateValueError

    def _validate_value_str(self):
        pass

    def _validate_arr(self):
        try:
            arr = json.loads(self.value)
        except ValueError:
            raise ValidateValueError

        if not isinstance(arr, list):
            raise ValidateValueError
        return arr

    def _validate_arr_each_type(self, types):
        arr = self._validate_arr()
        for item in arr:
            if type(item) not in types:
                raise ValidateValueError

    def _validate_value_arrInt(self):
        self._validate_arr_each_type(type_int)

    def _validate_value_arrFloat(self):
        self._validate_arr_each_type(type_float)

    def _validate_value_arrStr(self):
        self._validate_arr_each_type(type_str)

    def validate(self):
        self.validate_data_type()
        self.validate_value()
