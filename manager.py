# coding=utf-8
from functools import wraps
import os

import click
from tools.config_parser import ConfigParser
from tools.excel_to_json import ExcelToJson
from tools.excel_to_binary_stream import ExcelToBinaryStream
from tools.excel_validator import ExcelValidator
from tools.sheet_repeat_validator import SheetRepeatValidator


try:
    config = ConfigParser('config.local')
except IOError:
    config = ConfigParser('config.cfg')


def check_error(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        with open(config['ERROR_FILE_PATH']) as f:
            error = f.read()

        if error:
            return
        else:
            result = func(*args, **kwargs)
            print
            return result
    return wrapper


def yield_excel_path():
    for filename in os.listdir(config['EXCEL_FOLDER']):
        if filename.startswith('~') or not filename.endswith('.xlsx'):
            continue

        path = os.path.join(config['EXCEL_FOLDER'], filename)
        yield path


def clear_folder(folder_path):
    for filename in os.listdir(folder_path):
        os.remove(os.path.join(folder_path, filename))


@click.group()
def commands():
    pass


@commands.command()
@check_error
def export_json():
    print u'开始执行导出JSON'
    clear_folder(config['JSON_FOLDER'])
    for path in yield_excel_path():
        print u'导出JSON，正在处理:', path
        obj = ExcelToJson(path, config['JSON_FOLDER'])
        obj.export_data()


@commands.command()
@check_error
def export_csharp():
    print u'开始执行导出CSharp代码'
    clear_folder(config['BYTES_FOLDER'])
    clear_folder(config['CSHARP_CONFIG_FOLDER'])
    for path in yield_excel_path():
        print u'导出CSharp代码，正在处理:', path
        obj = ExcelToBinaryStream(path, config['BYTES_FOLDER'], config['CSHARP_CONFIG_FOLDER'])
        obj.export_data()


@commands.command()
def validate_excel():
    error_messages = []
    print u'开始执行excel校验'

    sheet_repeat_validator = SheetRepeatValidator()

    for path in yield_excel_path():
        print u'校验excel，正在处理', path
        obj = ExcelValidator(path)
        obj.run()
        error_messages.extend(obj.error_msgs)

        for sheet_name in obj.sheet_names:
            sheet_repeat_validator.add_name(sheet_name, obj.filename)

    sheet_repeat_validator.check_repeat()
    error_messages.extend(sheet_repeat_validator.error_messages)

    with open(config['ERROR_FILE_PATH'], 'w') as f:
        print '\n'.join(error_messages)
        f.write('\n'.join(error_messages).encode('utf-8'))

    if error_messages:
        tip = u'发生致命错误导致程序终止, 请检查错误日志文件 [%s]' % config['ERROR_FILE_PATH']
        print tip
        raw_input()


if __name__ == '__main__':
    commands()
