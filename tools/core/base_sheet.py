# coding=utf-8
from tools import const
from tools.models.grid import Grid


class BaseSheet(object):
    def __init__(self, sheet):
        self.sheet = sheet
        self.grid_list = []
        self.raw_values = [x for x in self.sheet.values]
        self.name = self.sheet.title

        self.descriptions = []
        self.comments = []
        self.keys = []
        self.scopes = []
        self.raw_data_types = []

    @property
    def primary_keys(self):
        """
        获取所有主键
        :return:
        """
        return [x[0].value for x in self.grid_list]

    def __repr__(self):
        return '<Sheet: %s>' % self.name

    def has_scope(self, scope):
        for item in self.scopes:
            if scope in item:
                return True

    def get_grid_by_scope(self, scope):
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
        types = []
        for data_type in self.raw_data_types:
            types.append(const.BASE_DATA_TYPES[data_type])
        return types

    def handle(self):
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
