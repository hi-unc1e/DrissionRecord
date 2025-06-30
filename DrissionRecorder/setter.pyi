# -*- coding:utf-8 -*-
from pathlib import Path
from typing import Union, Any, Optional, List, Tuple, Dict

from .base import OriginalRecorder, BaseRecorder
from .cell_style import CellStyle
from .db_recorder import DBRecorder
from .recorder import Recorder
from .tools import Header


class OriginalSetter(object):
    _recorder: OriginalRecorder = ...

    def __init__(self, recorder: OriginalRecorder): ...

    def cache_size(self, size: int) -> OriginalSetter:
        """设置缓存大小
        :param size: 缓存大小
        :return: 设置对象自己
        """
        ...

    def path(self, path: Union[str, Path]) -> OriginalSetter:
        """设置文件路径
        :param path: 文件路径
        :return: 设置对象自己
        """
        ...

    def show_msg(self, on_off: bool) -> OriginalSetter:
        """设置是否显示运行信息
        :param on_off: bool表示开关
        :return: 设置对象自己
        """
        ...

    def auto_backup(self,
                    interval: int = None,
                    path: Union[str, Path] = None,
                    new_name: bool = None) -> OriginalSetter:
        """设置自动备份相关参数
        :param interval: 自动保存多少次时触发备份，为0表示不自动备份，为None时不修改已设置值（初始为0）
        :param path: 备份文件存放路径，为None时不修改已设置值（初始为 'backup'）
        :param new_name: 是否用新名称备份，为None时不修改已设置值（初始为True）
        :return: 设置对象自己
        """
        ...


class BaseSetter(OriginalSetter):
    _recorder: BaseRecorder = ...

    def table(self, name: Union[str, bool]) -> BaseSetter:
        """设置默认表名
        :param name: 表名
        :return: 设置对象自己
        """
        ...

    def auto_new_header(self, on_off: bool = True) -> BaseSetter:
        """数据中有表头不存在的列时是否自动添加到表头，只有xlsx和csv格式有效
        :param on_off: bool表示开关
        :return: 设置对象自己
        """
        ...

    def before(self, data: Any) -> BaseSetter:
        """设置在数据前面补充的列
        :param data: 列表、元组或字符串，为字符串时则补充一列
        :return: 设置对象自己
        """
        ...

    def after(self, data: Any) -> BaseSetter:
        """设置在数据后面补充的列
        :param data: 列表、元组或字符串，为字符串时则补充一列
        :return: 设置对象自己
        """
        ...

    def _set_after_before(self, before: bool, data: Any) -> BaseSetter: ...


class RecorderSetter(BaseSetter):
    _recorder: Recorder = ...

    def __init__(self, recorder: Recorder): ...

    # -------------------上级开始-------------------
    def cache_size(self, size: int) -> RecorderSetter:
        """设置缓存大小
        :param size: 缓存大小
        :return: 设置对象自己
        """
        ...

    def show_msg(self, on_off: bool) -> RecorderSetter:
        """设置是否显示运行信息
        :param on_off: bool表示开关
        :return: 设置对象自己
        """
        ...

    def auto_backup(self,
                    interval: int = None,
                    path: Union[str, Path] = None,
                    new_name: bool = None) -> RecorderSetter:
        """设置自动备份相关参数
        :param interval: 自动保存多少次时触发备份，为0表示不自动备份，为None时不修改已设置值（初始为0）
        :param path: 备份文件存放路径，为None时不修改已设置值（初始为 'backup'）
        :param new_name: 是否用新名称备份，为None时不修改已设置值（初始为True）
        :return: 设置对象自己
        """
        ...

    def table(self, name: Union[str, bool]) -> RecorderSetter:
        """设置默认表名
        :param name: 表名
        :return: 设置对象自己
        """
        ...

    def auto_new_header(self, on_off: bool = True) -> RecorderSetter:
        """数据中有表头不存在的列时是否自动添加到表头，只有xlsx和csv格式有效
        :param on_off: bool表示开关
        :return: 设置对象自己
        """
        ...

    def after(self, data: Any) -> RecorderSetter:
        """设置在数据后面补充的列
        :param data: 列表、元组或字符串，为字符串时则补充一列
        :return: 设置对象自己
        """
        ...

    def before(self, data: Any) -> RecorderSetter:
        """设置在数据前面补充的列
        :param data: 列表、元组或字符串，为字符串时则补充一列
        :return: 设置对象自己
        """
        ...

    # -------------------上级结束-------------------

    def path(self, path: Union[str, Path], file_type: str = None) -> RecorderSetter:
        """设置文件路径
        :param path: 文件路径
        :param file_type: 指定文件类型
        :return: 设置对象自己
        """
        ...

    def header(self, header: Union[list, tuple], table: str = None,
               to_file: bool = True, row: int = None) -> RecorderSetter:
        """设置表头。只有 csv 和 xlsx 格式支持设置表头
        :param header: 表头，列表或元组
        :param table: 表名，只xlsx格式文件有效
        :param to_file: 是否写入到文件
        :param row: 指定写入文件的行号，不改变对象已设置的header_row属性，to_file为False时无效
        :return: 设置对象自己
        """
        ...

    def header_row(self, num: int, table=None) -> RecorderSetter:
        """设置标题行号
        :param num: 行号
        :param table: 表名，None为活动表格
        :return: 设置对象自己
        """
        ...

    def encoding(self, encoding: str) -> BaseSetter:
        """设置编码
        :param encoding: 编码格式
        :return: 设置对象自己
        """
        ...

    def delimiter(self, delimiter: str) -> RecorderSetter:
        """设置csv文件分隔符
        :param delimiter: 分隔符
        :return: 设置对象自己
        """
        ...

    def quote_char(self, quote_char: str) -> RecorderSetter:
        """设置csv文件引用符
        :param quote_char: 引用符
        :return: 设置对象自己
        """
        ...

    def follow_styles(self, on_off: bool = True) -> RecorderSetter:
        """设置是否跟随上一行的style，只有xlsx格式有效
        :param on_off: True或False
        :return: 设置对象自己
        """
        ...

    def file_type(self, file_type: str) -> RecorderSetter:
        """指定文件类型，无视文件后缀名
        :param file_type: 文件类型，可与路径后缀不一致
        :return: 设置对象自己
        """
        ...

    def new_row_height(self, height: float) -> RecorderSetter:
        """设置新行行高，只有xlsx格式有效
        :param height: 行高，传入None清空设置
        :return: 设置对象自己
        """
        ...

    def new_row_styles(self,
                       styles: Union[CellStyle, List[CellStyle], Tuple[CellStyle, ...],
                       Dict[Any, CellStyle], None]) -> RecorderSetter:
        """设置新行样式，只有xlsx格式有效，可传入多个，传入None则取消
        :param styles: 传入CellStyle对象设置整个新行，传入CellStyle对象组成的列表设置多个，传入None清空设置
        :return: 设置对象自己
        """
        ...

    def data_col(self, col: Union[None, str, int]) -> RecorderSetter:
        """设置用于填充数据的列
        :param col: 列号或列名，None表示header后面新列
        :return: 设置对象自己
        """
        ...

    def link_style(self, style: Union[CellStyle, True] = True) -> RecorderSetter:
        """设置单元格的链接样式
        :param style: CellStyle对象
        :return: 设置对象自己
        """
        ...


class DBSetter(BaseSetter):
    _recorder: DBRecorder = ...

    def __init__(self, recorder: DBRecorder): ...

    # -------------------上级开始-------------------
    def cache_size(self, size: int) -> DBSetter:
        """设置缓存大小
        :param size: 缓存大小
        :return: 设置对象自己
        """
        ...

    def show_msg(self, on_off: bool) -> DBSetter:
        """设置是否显示运行信息
        :param on_off: bool表示开关
        :return: 设置对象自己
        """
        ...

    def auto_backup(self,
                    interval: int = None,
                    path: Union[str, Path] = None,
                    new_name: bool = None) -> DBSetter:
        """设置自动备份相关参数
        :param interval: 自动保存多少次时触发备份，为0表示不自动备份，为None时不修改已设置值（初始为0）
        :param path: 备份文件存放路径，为None时不修改已设置值（初始为 'backup'）
        :param new_name: 是否用新名称备份，为None时不修改已设置值（初始为True）
        :return: 设置对象自己
        """
        ...

    def auto_new_header(self, on_off: bool = True) -> DBSetter:
        """数据中有表头不存在的列时是否自动添加到表头，只有xlsx和csv格式有效
        :param on_off: bool表示开关
        :return: 设置对象自己
        """
        ...

    def after(self, data: Any) -> DBSetter:
        """设置在数据后面补充的列
        :param data: 列表、元组或字符串，为字符串时则补充一列
        :return: 设置对象自己
        """
        ...

    def before(self, data: Any) -> DBSetter:
        """设置在数据前面补充的列
        :param data: 列表、元组或字符串，为字符串时则补充一列
        :return: 设置对象自己
        """
        ...
    # -------------------上级结束-------------------

    def path(self, path: Union[str, Path], table: Optional[str] = None) -> DBSetter:
        """重写父类方法
        :param path: 文件路径
        :param table: 数据表名称
        :return: 设置对象自己
        """
        ...

    def table(self, name: Union[str, bool]) -> DBSetter:
        """设置默认表名
        :param name: 表名
        :return: 设置对象自己
        """
        ...


def set_csv_header(recorder: Recorder,
                   header: Header,
                   row: int) -> None:
    """设置csv文件的表头
    :param recorder: Recorder对象
    :param header: 表头列表或元组
    :param row: 行号
    :return: None
    """
    ...


def set_xlsx_header(recorder: Recorder,
                    header: Header,
                    table: str,
                    row: int) -> None:
    """设置xlsx文件的表头
    :param recorder: Recorder对象
    :param header: 表头列表或元组
    :param table: 工作表名称
    :param row: 行号
    :return: None
    """
    ...
