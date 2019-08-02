# coding=utf-8
import os

import click
from tools.config_parser import ConfigParser
from tools.excel_to_json import ExcelToJson
from tools.excel_to_binary_stream import ExcelToBinaryStream


try:
    config = ConfigParser('config.local')
except IOError:
    config = ConfigParser('config.cfg')


def yield_excel_path():
    for filename in os.listdir(config['EXCEL_FOLDER']):
        if filename.startswith('~') or not filename.endswith('.xlsx'):
            continue

        path = os.path.join(config['EXCEL_FOLDER'], filename)
        yield path


@click.group()
def commands():
    pass


@commands.command()
def export_json():
    for path in yield_excel_path():
        obj = ExcelToJson(path, config['JSON_FOLDER'])
        obj.export_data()


@commands.command()
def export_csharp():
    for path in yield_excel_path():
        obj = ExcelToBinaryStream(path, config['BYTES_FOLDER'], config['CSHARP_CONFIG_FOLDER'])
        obj.export_data()


if __name__ == '__main__':
    commands()
