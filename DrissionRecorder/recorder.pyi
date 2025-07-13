# -*- coding:utf-8 -*-
"""
_header格式：{表名: Header对象}
表名为None表示默认sheet或csv文件
未设置header时值为None
"""
from csv import writer as csv_writer
from io import TextIOWrapper
from pathlib import Path
from typing import Any, Optional, Union, List, Dict, Tuple, Callable, Iterable

from openpyxl.worksheet.worksheet import Worksheet

from .base import BaseRecorder
from .setter import RecorderSetter
from .cell_style import CellStyle
from .tools import Header, RowData


class Recorder(BaseRecorder):
    _set: RecorderSetter = ...
    _row_height: Optional[float] = ...
    _follow_styles: bool = ...
    _styles: Union[CellStyle, List[CellStyle], Tuple[CellStyle], Dict[str, CellStyle], None] = ...
    _quote_char: str = ...
    _delimiter: str = ...
    _header: Dict[Optional[str], Optional[Header]] = ...
    _header_row: Dict[Optional[str], int] = ...
    _fast: bool = ...
    _methods: dict = ...
    _link_style: Optional[CellStyle] = ...
    _data: Dict[Optional[str], list] = ...
    _None_header_is_newest: Optional[bool] = ...
    _None_header_row_is_newest: Optional[bool] = ...
    data: Dict[Optional[str], list] = ...
    data_col: int = ...

    def __init__(self, path: Union[str, Path] = None, cache_size: int = 1000):
        """用于缓存并记录数据，可在达到一定数量时自动记录，以降低文件读写次数，减少开销
        :param path: 保存的文件路径
        :param cache_size: 每接收多少条记录写入文件，0为不自动写入
        """
        ...

    def _set_methods(self, file_type: str) -> None:
        """设置各种情况下使用的方法"""
        ...

    @property
    def set(self) -> RecorderSetter:
        """返回用于设置属性的对象"""
        ...

    @property
    def delimiter(self) -> str:
        """返回csv文件分隔符"""
        ...

    @property
    def quote_char(self) -> str:
        """返回csv文件引用符"""
        ...

    @property
    def header(self) -> Header:
        """返回表头，只支持csv和xlsx格式"""
        ...

    def add_data(self,
                 data: Any,
                 coord: Union[list, Tuple[Union[None, int, str], Union[None, int, str]], str, int] = None,
                 table: Union[str, bool] = None) -> None:
        """添加数据，可一次添加多条数据
        :param data: 插入的数据，任意格式，可以为二维数据
        :param coord: 要添加数据的坐标，可输入行号、列号或行列坐标，当格式不是xlsx或csv时无效，eg.'a3'、1、[3, 1]、'c'、'-3'
        :param table: 要写入的数据表，仅支持xlsx格式。为None表示用set.table()方法设置的值，为True表示活动的表格
        :return: None
        """
        ...

    def add_link(self,
                 link: Optional[str],
                 coord: Union[int, str, tuple],
                 content: Any = None,
                 table: Union[str, True, None] = None) -> None:
        """为单元格设置超链接
        :param link: 超链接，为None时删除链接
        :param coord: 单元格坐标
        :param content: 单元格内容
        :param table: 数据表名，仅支持xlsx格式。为None表示用set.table()方法设置的值，为Ture表示活动的表格
        :return: None
        """
        ...

    def add_img(self,
                img_path: Union[None, str, Path, dict],
                coord: Union[int, str, tuple],
                width: float = None,
                height: float = None,
                table: Union[str, True, None] = None) -> None:
        """向单元格设置图片
        :param img_path: 图片路径
        :param coord: 单元格坐标
        :param width: 图片宽
        :param height: 图片高
        :param table: 数据表名，仅支持xlsx格式。为None表示用set.table()方法设置的值，为Ture表示活动的表格
        :return: None
        """
        ...

    def add_styles(self,
                   styles: Union[CellStyle, dict, list, tuple],
                   coord: Union[int, str, tuple, list],
                   replace: bool = True,
                   table: Union[str, True, None] = None) -> None:
        """为单元格设置样式，可批量设置范围内的单元格
        :param styles: CellStyle对象，为None则清除单元格样式
        :param coord: 单元格坐标，输入数字可设置整行，输入列名字符串可设置整列，输入'A1:C5'、'a:d'、'1:5'格式可设置指定范围
        :param replace: 是否直接替换已有样式，运行效率较高，但不能单独修改某个属性
        :param table: 数据表名，仅支持xlsx格式。为None表示用set.table()方法设置的值，为Ture表示活动的表格
        :return: None
        """
        ...

    def add_rows_height(self, height: float, rows: Union[int, str, list, tuple, True],
                        table: Union[str, True, None] = None) -> None:
        """设置行高，可设置多行
        :param height: 行高
        :param rows: 行号，可指定多行（1、'1:4'、[1, 2, 3]），为Ture设置所有行
        :param table: 数据表名，仅支持xlsx格式。为None表示用set.table()方法设置的值，为Ture表示活动的表格
        :return: None
        """
        ...

    def add_cols_width(self, width: float, cols: Union[int, str, list, tuple, True],
                       table: Union[str, True, None] = None, is_header: bool = False) -> None:
        """设置列宽，可设置多列
        :param width: 列宽
        :param cols: 列号，可指定多列（1、'a'、'1:4'、'a:d'、[1, 2, 3]、['a', 'b', 'c']），为Ture设置所有列
        :param table: 数据表名，仅支持xlsx格式。为None表示用set.table()方法设置的值，为Ture表示活动的表格
        :param is_header: cols里的str是列号还是header
        :return: None
        """
        ...

    def _add_link(self,
                  coord: Union[int, str, tuple, list],
                  link: Union[str, dict],
                  content: Union[None, int, str, float] = None,
                  table: Union[str, bool] = None) -> None:
        """为单元格设置超链接
        :param coord: 单元格坐标
        :param link: 超链接，为None时删除链接
        :param content: 单元格内容
        :param table: 数据表名，仅支持xlsx格式。为None表示用set.table()方法设置的值，为bool表示活动的表格
        :return: None
        """
        ...

    def _add_img(self,
                 coord: Union[int, str, tuple, list],
                 img_path: Union[None, str, Path, dict],
                 width: float = None,
                 height: float = None,
                 table: Union[str, bool] = None) -> None:
        """向单元格设置图片
        :param coord: 单元格坐标
        :param img_path: 图片路径
        :param width: 图片宽
        :param height: 图片高
        :param table: 数据表名，仅支持xlsx格式。为None表示用set.table()方法设置的值，为bool表示活动的表格
        :return: None
        """
        ...

    def _add_styles(self,
                    coord: Union[int, str, tuple, list],
                    style: Union[CellStyle, dict],
                    replace: bool = True,
                    table: Union[str, bool] = None) -> None:
        """为单元格设置样式，可批量设置范围内的单元格
        :param coord: 单元格坐标，输入数字可设置整行，输入列名字符串可设置整列，输入'A1:C5'、'a:d'、'1:5'格式可设置指定范围
        :param style: CellStyle对象，为None则清除单元格样式
        :param replace: 是否直接替换已有样式，运行效率较高，但不能单独修改某个属性
        :param table: 数据表名，仅支持xlsx格式。为None表示用set.table()方法设置的值，为bool表示活动的表格
        :return: None
        """
        ...

    def _add_rows_height(self, rows: Union[int, str, list, tuple, True], height: float,
                         table: Union[str, bool] = None) -> None:
        """设置行高，可设置连续多行
        :param rows: 行号，可指定多行（1、'1:4'、[1, 2, 3]），为Ture设置所有行
        :param height: 行高
        :param table: 数据表名，仅支持xlsx格式。为None表示用set.table()方法设置的值，为bool表示活动的表格
        :return: None
        """
        ...

    def _add_cols_width(self, cols: Union[int, str, list, tuple, True], width: float,
                        table: Union[str, bool] = None, is_header: bool = False) -> None:
        """设置列宽，可设置多列
        :param cols: 列号，可指定多列（1、'a'、'序号'、'1:4'、'a:d'、[1, 2, 3]、['a', 'b', 'c']），为Ture设置所有列
        :param width: 列宽
        :param table: 数据表名，仅支持xlsx格式。为None表示用set.table()方法设置的值，为True表示活动的表格
        :param is_header: cols里的str是列号还是header
        :return: None
        """
        ...

    def _add(self, data: dict, table: Optional[str], to_slow: bool, num: int, add_method: callable) -> None:
        """为单元格设置样式，可批量设置范围内的单元格
        :param data: 数据，[dict, ...]格式
        :param table: 表格名称，None为活动表格
        :param to_slow: 是否转到slow模式
        :param num: 增加的数据量
        :return: None
        """
        ...

    def _add_data_any(self, data: dict, table: Optional[str]) -> None:
        """添加data到_data的操作，适用于任意类型文件
        :param data: 数据
        :param table: 要添加数据的表
        :return: None
        """
        ...

    def _add_data_txt(self, data: dict, table: Optional[str]) -> None:
        """添加data到_data的操作，适用于除xlsx以外的文件类型
        :param data: 数据
        :param table: 要添加数据的表
        :return: None
        """
        ...

    def _add_others(self, data: dict, table: Optional[str]) -> None:
        """添加style、link等到_data的操作
        :param data: 数据
        :param table: 要添加数据的表
        :return: None
        """
        ...

    def rows(self,
             key_cols: Union[str, int, list, tuple, bool] = True,
             sign_col: Union[str, int, True] = True,
             signs: Any = None,
             deny_sign: bool = False,
             count: int = None,
             begin_row: Optional[int] = None,
             end_row: Optional[int] = None) -> List[RowData]:
        """返回符合条件的行数据，可指定只要某些列
        :param key_cols: 作为关键字的列，可以是多列，传入表头值或列序号，要用列号用Col('a')，为True获取所有列
        :param sign_col: 用于筛选数据的列，传入表头值或列序号，要用列号用Col('a')，为True获取所有行
        :param signs: 按这个值筛选目标行，可用list, tuple, set设置多个
        :param deny_sign: 是否反向匹配sign，即筛选指不是sign的行
        :param count: 获取多少条数据，为None获取所有
        :param begin_row: 数据开始的行，None表示header_row后面一行
        :param end_row: 数据结束的行，None最后一行
        :return: RowData对象
        """
        ...

    def _handle_data(self, data: Any, coord: tuple) -> Tuple[dict, int]:
        """把数据处理成存储格式
        :param data: 要处理的数据
        :param coord: 单元格坐标
        :return: (处理后的数据, 数据长度)
        """
        ...

    def _record(self) -> None:
        """记录数据"""
        ...

    def _fast_mode(self) -> None:
        """切换到fast模式"""
        ...

    def _slow_mode(self) -> None:
        """切换到slow模式"""
        ...

    def _to_xlsx_fast(self) -> None:
        """fast模式填写数据到xlsx文件"""
        ...

    def _to_csv_fast(self) -> None:
        """fast模式填写数据到csv文件"""
        ...

    def _to_csv_slow(self) -> None:
        """slow模式填写数据到csv文件"""
        ...

    def _to_txt_fast(self) -> None:
        """记录数据到txt文件"""
        ...

    def _to_jsonl_fast(self) -> None:
        """记录数据到jsonl文件"""
        ...

    def _to_json_fast(self) -> None:
        """记录数据到json文件"""
        ...

    def _to_txt_slow(self) -> None:
        """记录数据到txt文件"""
        ...

    def _to_jsonl_slow(self) -> None:
        """记录数据到jsonl文件"""
        ...

    def _to_json_slow(self) -> None:
        """记录数据到json文件"""
        ...


def handle_txt_lines(data_lst: list, lines: list, val: Any, method: Callable) -> None:
    """txt、json、jsonl格式相同的写入逻辑
    :param data_lst: 数据总列表
    :param lines: readlines()从文件读取的原数据列表
    :param val: 插入空行时的值
    :param method: 处理单个数据使用的方法
    :return: None
    """
    ...


def handle_txt_data(lines: list, num: int, data: Union[dict, list]) -> None:
    """处理txt格式单个数据的方法，对应handle_txt_lines()的method
    :param lines: list格式的文件数据
    :param num: 行号
    :param data: 要写入的数据
    :return: None
    """
    ...


def handle_jsonl_data(lines: list, num: int, data: Union[dict, list]) -> None:
    """处理jsonl格式单个数据的方法，对应handle_txt_lines()的method
    :param lines: list格式的文件数据
    :param num: 行号
    :param data: 要写入的数据
    :return: None
    """
    ...


def handle_json_data(lines: list, num: int, data: Union[dict, list]) -> None:
    """处理json格式单个数据的方法，对应handle_txt_lines()的method
    :param lines: list格式的文件数据
    :param num: 行号
    :param data: 要写入的数据
    :return: None
    """
    ...


def get_header(recorder: Recorder, ws: Worksheet = None) -> Header:
    """获取当前指定的table的header
    :param recorder: Recorder对象
    :param ws: Worksheet对象
    :return: Header对象
    """
    ...


def handle_new_sheet(recorder: Recorder, ws: Worksheet, data: list) -> int:
    """从设置或第一条dict数据获取表头并向新表写入
    :param recorder: Recorder对象
    :param ws: 数据表对象
    :param data: 对应数据表的数据列表
    :return: 开始写数据的行的前一行
    """
    ...


def get_first_dict(data: list) -> dict:
    """判断数据集第一条是否dict，如果第一条是二维数据，判断其第一条是否dict，是则返回它
    :param data: 数据列表
    :return: 第一条dict格式数据
    """
    ...


def get_xlsx_rows(recorder: Recorder, header: Header, key_cols: Union[list, True],
                  begin_row: Optional[int], end_row: Optional[int],
                  sign_col: Union[str, int, bool], sign: Any,
                  deny_sign: bool, count: int, ws: Worksheet) -> List[RowData]:
    """获取xlsx文件指定行数据
    :param recorder: Recorder对象
    :param header: Header对象
    :param key_cols: 要获取的列，为True获取所有，可指定多列
    :param begin_row: 开始行号
    :param end_row: 结束行号，None为最后一行
    :param sign_col: 作为条件的列
    :param sign: 按这个值筛选目标行，可设置多个
    :param deny_sign: 是否反向匹配sign，即筛选指不是sign的行
    :param count: 获取多少条数据，为None获取所有
    :param ws: Worksheet对象
    :return: 获取到的数据列表
    """
    ...


def get_csv_rows(recorder: Recorder, header: Header, key_cols: Union[list, True],
                 begin_row: Optional[int], end_row: Optional[int],
                 sign_col: Union[str, int, bool], sign: Any,
                 deny_sign: bool, count: int, ws: Worksheet) -> List[RowData]:
    """获取csv文件指定行数据
    :param recorder: Recorder对象
    :param header: Header对象
    :param key_cols: 要获取的列，为True获取所有，可指定多列
    :param begin_row: 开始行号
    :param end_row: 结束行号，None为最后一行
    :param sign_col: 作为条件的列
    :param sign: 按这个值筛选目标行，可设置多个
    :param deny_sign: 是否反向匹配sign，即筛选指不是sign的行
    :param count: 获取多少条数据，为None获取所有
    :param ws: 与get_xlsx_rows()对应
    :return: 获取到的数据列表
    """
    ...


def get_xlsx_rows_with_count(key_cols: Union[list, True], deny_sign: bool, header: Header, rows: Iterable,
                             begin_row: Optional[int], end_row: Optional[int],
                             sign_col: Union[str, int, bool], sign: Any,
                             count: int) -> List[RowData]:
    """执行从xlsx中获取数据，有指定数量
    :param key_cols: 要获取的列，True为所有
    :param deny_sign: 是否反向匹配sign，即筛选指不是sign的行
    :param header: Header对象
    :param rows: 行组成的列表
    :param begin_row: 开始行号
    :param end_row: 结束行号，None为最后一行
    :param sign_col: 用于筛选数据的列
    :param sign: 用于筛选数据的值
    :param count: 数据总条数
    :return: 数据对象列表
    """
    ...


def get_xlsx_rows_without_count(key_cols: Union[list, True], deny_sign: bool, header: Header, rows: Iterable,
                                begin_row: Optional[int], end_row: Optional[int],
                                sign_col: Union[str, int, bool], sign: Any) -> List[RowData]:
    """执行从xlsx中获取全部数据
    :param key_cols: 要获取的列，True为所有
    :param deny_sign: 是否反向匹配sign，即筛选指不是sign的行
    :param header: Header对象
    :param rows: 行组成的列表
    :param begin_row: 开始行号
    :param end_row: 结束行号，None为最后一行
    :param sign_col: 用于筛选数据的列
    :param sign: 用于筛选数据的值
    :return: 数据对象列表
    """
    ...


def get_csv_rows_with_count(lines, begin_row: Optional[int], end_row: Optional[int],
                            sign_col: Union[str, int, bool], sign: Any, deny_sign: bool,
                            key_cols: Union[list, True], res, header: Header, count: int) -> List[RowData]:
    """执行从csv中获取数据，有指定数量
    :param lines:
    :param begin_row: 开始行号
    :param end_row: 结束行号，None为最后一行
    :param sign_col: 用于筛选数据的列
    :param sign: 用于筛选数据的值
    :param deny_sign: 是否反向匹配sign，即筛选指不是sign的行
    :param key_cols: 要获取的列，True为所有
    :param res: 结果列表
    :param header: Header对象
    :param count: 数据总条数
    :return: 数据对象列表
    """
    ...


def get_csv_rows_without_count(lines, begin_row: Optional[int], end_row: Optional[int],
                               sign_col: Union[str, int, bool], sign: Any,
                               deny_sign: bool, key_cols: Union[list, True], res, header: Header) -> List[RowData]:
    """执行从csv中获取全部数据
    :param lines:
    :param begin_row: 开始行号
    :param end_row: 结束行号，None为最后一行
    :param sign_col: 用于筛选数据的列
    :param sign: 用于筛选数据的值
    :param deny_sign: 是否反向匹配sign，即筛选指不是sign的行
    :param key_cols: 要获取的列，True为所有
    :param res: 结果列表
    :param header: Header对象
    :return: 数据对象列表
    """
    ...


def get_and_set_csv_header(recorder: Recorder, new_csv: bool, file: TextIOWrapper, writer: csv_writer) -> None:
    """从csv获取表头或把已获取的表头设置到新csv
    :param recorder: Recorder对象
    :param new_csv: 是否新csv文件
    :param file: 文件对象
    :param writer: csv writer对象
    :return: None
    """
    ...
