# -*- coding:utf-8 -*-
from pathlib import Path
from sqlite3 import Connection, Cursor
from typing import Union, Any, Optional

from .base import BaseRecorder
from .setter import DBSetter


class DBRecorder(BaseRecorder):
    _conn: Optional[Connection] = ...
    _cur: Optional[Cursor] = ...
    _setter: Optional[DBSetter] = ...
    _data: dict = ...
    data: dict = ...

    def __init__(self,
                 path: Union[str, Path] = None,
                 cache_size: int = 1000,
                 table: str = None):
        """用于存储数据到sqlite的工具
        :param path: 保存的文件路径
        :param cache_size: 每接收多少条记录写入文件，0为不自动写入
        :param table: 默认表名
        """
        ...

    @property
    def set(self) -> DBSetter:
        """返回用于设置属性的对象"""
        ...

    @property
    def tables(self) -> list:
        """返回所有表名"""
        ...

    def add_data(self, data: Any, table: str = None) -> None:
        """添加数据
        :param data: 可以是一维或二维数据，dict格式可向对应列填写数据，其余格式按顺序从左到右填入各列
        :param table: 数据要插入的表名称
        :return: None
        """
        ...

    def run_sql(self, sql: str, single: bool = True, commit: bool = False) -> Union[None, list, tuple]:
        """执行sql语句并返回结果
        :param sql: sql语句
        :param single: 是否只获取一个结果
        :param commit: 是否提交到数据库
        :return: 查找到的结果，没有结果时返回None
        """
        ...

    def _connect(self) -> None:
        """连接数据库"""
        ...

    def _close_connection(self) -> None:
        """关闭数据库 """
        ...

    def _record(self) -> None:
        """保存数据到sqlite"""
        ...

    def _to_database(self,
                     data_list: list,
                     table: str,
                     tables: dict) -> None:
        """把数据批量写入指定数据表
        :param data_list: 要写入的数据组成的列表
        :param table: 要写入数据的数据表名称
        :param tables: 数据库中数据表和列信息
        :return: None
        """
        ...

    def _handle_data(self, data: Any) -> list:
        """接收数据后的格式化"""
        ...
