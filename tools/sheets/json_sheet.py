# coding=utf-8
import os
import json
from tools.core.base_sheet import BaseSheet


class JsonSheet(BaseSheet):
    def __init__(self, path):
        super(JsonSheet, self).__init__(path)

    def get_json_data(self):
        data_list = []
        for row in self.grid_list:
            data = dict()
            for grid in row:
                data[grid.key] = grid.value
            data_list.append(data)
        return data_list

    def save_file(self, export_dir_path):
        filename = '%s.json' % self.name
        path = os.path.join(export_dir_path, filename)
        data = self.get_json_data()
        text = json.dumps(data, ensure_ascii=False, indent=4)
        with open(path, 'w') as f:
            f.write(text.encode('utf-8'))
