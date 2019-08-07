# coding=utf-8
import struct


class StreamTool(object):
    def __init__(self, head='<'):
        self.buffs = []
        self.fmt = [head]

    def write_int(self, data):
        self.fmt.append('i')
        self.buffs.append(data)

    def write_float(self, data):
        self.fmt.append('f')
        self.buffs.append(data)

    def write_str(self, data):
        data = data.encode('utf-8')
        self.fmt.append('i')
        str_len = max(1, len(data))
        self.fmt.append('%ss' % str_len)
        self.buffs.append(str_len)
        self.buffs.append(data)

    def write_arrInt(self, data):
        array_len = len(data)
        self.fmt.append('i')

        if array_len > 0:
            self.fmt.append('%si' % array_len)

        self.buffs.append(array_len)
        for item in data:
            self.buffs.append(item)

    def write_arrFloat(self, data):
        array_len = len(data)
        self.fmt.append('i')
        if array_len > 0:
            self.fmt.append('%sf' % array_len)

        self.buffs.append(array_len)
        for item in data:
            self.buffs.append(item)

    def write_arrStr(self, data):
        self.fmt.append('i')
        self.buffs.append(len(data))
        for item in data:
            item = item.encode('utf-8')
            item_len = len(item)
            self.fmt.append('i')
            self.fmt.append('%ss' % item_len)
            self.buffs.append(item_len)
            self.buffs.append(item)

    def get_bytes(self):
        fmt = ''.join(self.fmt)
        binary_stream = struct.pack(fmt, *self.buffs)
        return binary_stream


if __name__ == '__main__':
    test = StreamTool()
    test.write_str('')
    print test.buffs
