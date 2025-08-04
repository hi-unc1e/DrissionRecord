# -*- coding:utf-8 -*-
from pathlib import Path
from time import sleep

from .base import OriginalRecorder


class ByteRecorder(OriginalRecorder):
    __END = (0, 2)

    def __init__(self, path=None, cache_size=1000):
        super().__init__(path, cache_size)
        self._type = 'byte'

    def add_data(self, data, seek=None):
        while self._pause_add:  # 等待其它线程写入结束
            sleep(.02)

        if not isinstance(data, bytes):
            raise TypeError('只能接受bytes类型数据。')
        if seek is not None and not (isinstance(seek, int) and seek >= 0):
            raise ValueError('seek参数只能接受None或大于等于0的整数。')

        self._data.append((data, seek))
        self._data_count += 1

        if 0 < self.cache_size <= self._data_count:
            self.record()

    def _record(self):
        if not self._file_exists and not Path(self.path).exists():
            with open(self.path, 'wb'):
                pass

        with open(self.path, 'rb+') as f:
            previous = None
            for i in self._data:
                loc = ByteRecorder.__END if i[1] is None else (i[1], 0)
                if not (previous == loc == ByteRecorder.__END):
                    f.seek(loc[0], loc[1])
                    previous = loc
                f.write(i[0])
