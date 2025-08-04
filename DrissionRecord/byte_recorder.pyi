# -*- coding:utf-8 -*-
from pathlib import Path
from typing import Union, Optional

from .base import OriginalRecorder


class ByteRecorder(OriginalRecorder):
    __END: tuple = ...
    data: list = ...

    def __init__(self,
                 path: Union[None, str, Path] = None,
                 cache_size: int = 1000):
        """用于记录字节数据的工具
        :param path: 保存的文件路径
        :param cache_size: 每接收多少条记录写入文件，0为不自动写入
        """
        ...

    def add_data(self,
                 data: bytes,
                 seek: int = None) -> None:
        """添加一段二进制数据
        :param data: bytes类型数据
        :param seek: 在文件中的位置，None表示最后
        :return: None
        """
        ...

    def _record(self) -> None:
        """记录数据到文件"""
        ...
