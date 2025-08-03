# -*- coding:utf-8 -*-
from abc import abstractmethod
from pathlib import Path
from threading import Lock
from time import sleep

from .setter import OriginalSetter, BaseSetter
from .tools import get_usable_path, make_valid_name, get_tables, make_final_data_simplify


class OriginalRecorder(object):
    _SUPPORTS = ('any',)

    def __init__(self, path=None, cache_size=1000):
        self._data = []
        self._path = None
        self._type = None
        self._lock = Lock()
        self._pause_add = False  # 文件写入时暂停接收输入
        self._pause_write = False  # 标记文件正在被一个线程写入
        self._setter = None
        self._data_count = 0  # 已缓存数据的条数
        self._file_exists = False
        self._backup_path = 'backup'
        self._backup_times = 0
        self._backup_interval = 0  # 多少次就自动保存
        self._backup_overwrite = False
        self.show_msg = True
        if path:
            self.set.path(path)
        self._cache = cache_size or 0

    def __del__(self):
        self.record()

    @property
    def set(self):
        if self._setter is None:
            self._setter = OriginalSetter(self)
        return self._setter

    @property
    def cache_size(self):
        return self._cache

    @property
    def path(self):
        return self._path

    @property
    def type(self):
        return self._type

    @property
    def data(self):
        return self._data

    def record(self):
        if not self._data_count:
            return self._path
        if not self._path:
            raise ValueError('保存路径为空。')

        with self._lock:
            if self._backup_interval and self._backup_times >= self._backup_interval:
                self.backup(folder=self._backup_path, overwrite=self._backup_overwrite)

            self._pause_add = True  # 写入文件前暂缓接收数据
            if self.show_msg:
                print(f'{self.path} 开始写入文件，切勿关闭进程。')

            Path(self.path).parent.mkdir(parents=True, exist_ok=True)
            while True:
                try:
                    while self._pause_write:  # 等待其它线程写入结束
                        sleep(.02)

                    self._pause_write = True
                    self._record()
                    break

                except PermissionError:
                    if self.show_msg:
                        print('\r文件被打开，保存失败，请关闭，程序会自动重试。', end='')

                except Exception as e:
                    try:
                        with open('failed_data.txt', 'a+', encoding='utf-8') as f:
                            f.write(str(self.data) + '\n')
                        print('保存失败的数据已保存到failed_data.txt。')
                    except:
                        raise e
                    raise

                finally:
                    self._pause_write = False

                sleep(.3)

            if self.show_msg:
                print(f'{self.path} 写入文件结束。')
            self.clear()
            self._pause_add = False

        if self._backup_interval:
            self._backup_times += 1
        self._file_exists = True
        return self._path

    def clear(self):
        self._data.clear()
        self._data_count = 0

    def backup(self, folder=None, name=None, overwrite=None):
        src_path = Path(self._path)
        if not self._file_exists:
            if not src_path.exists():
                return ''
            self._file_exists = True

        if overwrite is None:
            overwrite = self._backup_overwrite
        folder = Path(folder if folder else self._backup_path)
        folder.mkdir(parents=True, exist_ok=True)
        if not name:
            name = src_path.name
        elif not name.endswith(src_path.suffix):
            name = f'{name}{src_path.suffix}'
        path = folder / make_valid_name(name)
        if not overwrite and path.exists():
            from datetime import datetime
            name = f'{path.stem}_{datetime.now().strftime("%Y%m%d%H%M%S")}{path.suffix}'
            path = get_usable_path(folder / name)

        from shutil import copy
        copy(self._path, path)
        self._backup_times = 0
        return str(path.absolute())

    @abstractmethod
    def add_data(self, data):
        pass

    @abstractmethod
    def _record(self):
        pass


class BaseRecorder(OriginalRecorder):
    _SUPPORTS = ('xlsx', 'csv')

    def __init__(self, path=None, cache_size=None):
        super().__init__(path, cache_size)
        self._before = []
        self._after = []
        self._encoding = 'utf-8'
        self._table = None
        self._make_final_data = make_final_data_simplify
        self._auto_new_header = False

    @property
    def set(self):
        if self._setter is None:
            self._setter = BaseSetter(self)
        return self._setter

    @property
    def before(self):
        return self._before

    @property
    def after(self):
        return self._after

    @property
    def table(self):
        return self._table

    @property
    def tables(self):
        if self._type != 'xlsx':
            raise TypeError('只有xlsx格式能使用tables属性。')
        if not self._path:
            raise RuntimeError('未指定文件路径。')
        return get_tables(self._path)

    @property
    def encoding(self):
        return self._encoding

    @abstractmethod
    def add_data(self, data, table=None):
        pass

    @abstractmethod
    def _record(self):
        pass
