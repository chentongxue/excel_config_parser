# coding=utf-8
import os

import click
from tools.config_parser import ConfigParser
from tools.excel_to_json import ExcelToJson
from tools.excel_to_binary_stream import ExcelToBinaryStream
from tools.excel_validator import ExcelValidator


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


@commands.command()
def validate_excel():
    error_messages = []
    for path in yield_excel_path():
        obj = ExcelValidator(path)
        obj.run()
        error_messages.extend(obj.error_msgs)

    if error_messages:
        with open(config['ERROR_FILE_PATH'], 'w') as f:
            print '\n'.join(error_messages)
            f.write('\n'.join(error_messages).encode('utf-8'))


if __name__ == '__main__':
    commands()
