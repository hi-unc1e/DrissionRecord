# -*- coding:utf-8 -*-
from abc import abstractmethod
from pathlib import Path
from threading import Lock
from typing import Union, Optional, Callable

from .setter import OriginalSetter, BaseSetter


class OriginalRecorder(object):
    """记录器的基类"""
    _cache: int = ...
    _path: Optional[str] = ...
    _type: Optional[str] = ...
    _data: list = ...
    _lock: Lock = ...
    _pause_add: bool = ...
    _pause_write: bool = ...
    _setter: Optional[OriginalSetter] = ...
    _data_count: int = ...
    _file_exists: bool = ...
    _backup_path: str = ...
    _backup_times: int = ...
    _backup_interval: int = ...
    _backup_overwrite: bool = ...
    show_msg: bool = ...

    def __init__(self,
                 path: Union[str, Path, None] = None,
                 cache_size: int = None) -> None:
        """
        :param path: 保存的文件路径
        :param cache_size: 每接收多少条记录写入文件，0为不自动写入
        """
        ...

    def __del__(self) -> None:
        """对象关闭时把剩下的数据写入文件"""
        ...

    @property
    def set(self) -> OriginalSetter:
        """返回用于设置属性的对象"""
        ...

    @property
    def cache_size(self) -> int:
        """返回缓存大小"""
        ...

    @property
    def path(self) -> str:
        """返回文件路径"""
        ...

    @property
    def type(self) -> str:
        """返回文件类型"""
        ...

    @property
    def data(self) -> Union[dict, list]:
        """返回当前保存在缓存的数据"""
        ...

    def record(self) -> str:
        """记录数据，返回文件路径"""
        ...

    def clear(self) -> None:
        """清空缓存中的数据"""
        ...

    def backup(self,
               folder: Union[str, Path, None] = None,
               name: str = None,
               overwrite: bool = None) -> str:
        """把当前文件备份到指定路径
        :param folder: 文件夹路径，为None使用内置路径（初始 'backup'）
        :param name: 保存的文件名，可不含后缀，为None使用内置路径文件名
        :param overwrite: 是否覆盖同名文件，为False时每次备份文件名添加当前时间，为None使用内置设置
        """
        ...

    def delete(self) -> None:
        """删除所指向的文件"""
        ...

    @abstractmethod
    def add_data(self, data):
        ...

    @abstractmethod
    def _record(self):
        ...


class BaseRecorder(OriginalRecorder):
    """Recorder和DBRecorder的父类"""
    _encoding: str = ...
    _before: Optional[list] = ...
    _after: Optional[list] = ...
    _table: Optional[str] = ...
    _setter: BaseSetter = ...
    _make_final_data: Callable = ...
    _auto_new_header: bool = ...

    def __init__(self, path: Union[None, str, Path] = None, cache_size: int = None) -> None:
        """
        :param path: 保存的文件路径
        :param cache_size: 每接收多少条记录写入文件，0为不自动写入
        """
        ...

    @property
    def set(self) -> BaseSetter:
        """返回用于设置属性的对象"""
        ...

    @property
    def before(self) -> Union[dict, list]:
        """返回当前before内容"""
        ...

    @property
    def after(self) -> Union[dict, list]:
        """返回当前after内容"""
        ...

    @property
    def table(self) -> Optional[str]:
        """返回当前使用的表名"""
        ...

    @property
    def tables(self) -> list:
        """返回所有表名"""
        ...

    @property
    def encoding(self) -> str:
        """返回csv文件使用的编码格式"""
        ...

    @abstractmethod
    def _record(self):
        ...
