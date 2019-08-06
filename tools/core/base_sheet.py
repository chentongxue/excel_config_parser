# coding=utf-8
from tools import const
from tools.models.grid import Grid


class BaseSheet(object):
    """
    基础的 Sheet 类
    具体业务的 Sheet 处理，可继承该类编写具体逻辑
    """
    def __init__(self, sheet):
        """
        传入的 openpyxl sheet 对象
        :param sheet:
        """
        self.sheet = sheet
        self.grid_list = []  # 二维数组，用于存放 Grid 实例，第1维对应excel中的一行，第2维对应excel一行中的一个单元格
        self.raw_values = [x for x in self.sheet.values]  # 整个sheet的原始单元格数据
        self.name = self.sheet.title  # sheet 标题

        self.descriptions = []
        self.comments = []
        self.keys = []
        self.scopes = []
        self.raw_data_types = []

    @property
    def primary_keys(self):
        """
        获取所有主键的值，即获取第一列的值
        :return: [primary_key]
        """
        return [x[0].value for x in self.grid_list]

    def __repr__(self):
        return '<Sheet: %s>' % self.name

    def has_scope(self, scope):
        """是否含有某个作用域 (C/S)"""
        for item in self.scopes:
            if scope in item:
                return True

    def get_grid_by_scope(self, scope):
        """
        根据作用域(C/S)获取 Grid 实例二维对象列表
        :param scope:
        :return: [[Grid]]
        """

        if not self.has_scope(scope):
            return []

        grid_list = []
        for row in self.grid_list:
            row_list = []
            for grid in row:
                if scope in grid.scope:
                    row_list.append(grid)
            grid_list.append(row_list)
        return grid_list

    @property
    def data_types(self):
        """
        返回 sheet 的所有数据类型的值
        :return: list[int]
        """
        types = []
        for data_type in self.raw_data_types:
            types.append(const.BASE_DATA_TYPES[data_type])
        return types

    def handle(self):
        """
        处理 sheet 具体实现
        将所有数据单元格处理成 Grid 实例，以二维数组形式存入 self.grid_list 中
        :return:
        """
        self.descriptions = self.raw_values[const.DESCRIPTION_ROW_NUM]
        self.comments = self.raw_values[const.COMMENT_ROW_NUM]
        self.keys = self.raw_values[const.KEY_ROW_NUM]
        self.scopes = self.raw_values[const.SCOPE_ROW_NUM]
        self.raw_data_types = self.raw_values[const.DATA_TYPE_ROW_NUM]

        for row_num in xrange(const.DATA_START_ROW_NUM, len(self.raw_values)):
            row_grid_list = []
            row_values = self.raw_values[row_num]

            # 空行不处理
            if row_values[0] is None:
                continue

            # 以 // 开头的为注释行，不处理
            if isinstance(row_values[0], unicode) and row_values[0].startswith('//'):
                continue

            for col_num, cell_value in enumerate(row_values):
                if self.scopes[col_num] is None:
                    continue

                row_grid_list.append(Grid(
                    key=self.keys[col_num],
                    value=cell_value,
                    description=self.descriptions[col_num],
                    comment=self.comments[col_num],
                    scope=self.scopes[col_num],
                    data_type=self.raw_data_types[col_num],
                    row_num=row_num,
                    col_num=col_num
                ))
            self.grid_list.append(row_grid_list)
