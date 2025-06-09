# -*- coding:utf-8 -*-
from io import TextIOWrapper
from pathlib import Path
from typing import Union, Tuple, Any, Optional, List, Dict, Iterable

from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from .base import BaseRecorder
from .recorder import Recorder


def remove_end_Nones(in_list: list) -> list:
    """去除列表后面所有None
    :param in_list: 要处理的list
    """
    ...


class BaseHeader(object):
    _NUM_KEY: dict = ...
    _KEY_NUM: dict = ...
    _CONTENT_FUNCS: dict = ...

    @property
    def key_num(self) -> Dict[str, int]: ...

    @property
    def num_key(self) -> Dict[int, str]: ...

    def __iter__(self): ...


class Header(BaseHeader):

    def __init__(self, header: Union[list, tuple] = None): ...

    def __getitem__(self, item: Union[int, str]): ...

    def __len__(self) -> int: ...

    def values(self): ...

    def items(self): ...

    def make_row_data(self, row: int, row_values: dict, None_val: Optional[''] = None) -> RowData:
        """
        :param row: 行号
        :param row_values: {列序号: 值}
        :param None_val: 空值是None还是''
        :return: RowData对象
        """
        ...

    def make_insert_list(self, data, file_type: Optional[str], rewrite: bool) -> Tuple[list, bool]:
        """生成写入文件list格式的新行数据
        :param data: 待处理行数据
        :param file_type: 文件类型，用于选择处理方法
        :param rewrite: 只用于对齐参数
        :return: 处理后的行数据
        """
        ...

    def make_change_list(self, line_data, data, col: int,
                         file_type: Optional[str], rewrite: bool) -> Tuple[list, bool]:
        """生产写入文件list格式的原有行数据
        :param line_data: 原有行数据
        :param data: 待处理行数据
        :param col: 要写入的列
        :param file_type: 文件类型，用于选择处理方法
        :param rewrite: 只用于对齐参数
        :return: (处理后的行数据, 是否重写表头)
        """
        ...

    def make_insert_list_rewrite(self, data, file_type: Optional[str], rewrite: bool) -> Tuple[list, bool]:
        """生产写入文件list格式的新行数据
        :param data: 待处理行数据
        :param rewrite: 是否需要重写表头
        :param file_type: 文件类型，用于选择处理方法
        :return: (处理后的行数据, 是否重写表头)
        """
        ...

    def make_change_list_rewrite(self, line_data, data, col: int, file_type, rewrite: bool) -> Tuple[list, bool]:
        """生产写入文件list格式的原有行数据
        :param line_data: 原有行数据
        :param data: 待处理行数据
        :param col: 要写入的列
        :param rewrite: 是否需要重写表头
        :param file_type: 文件类型，用于选择处理方法
        :return: (处理后的行数据, 是否重写表头)
        """
        ...

    def make_num_dict(self, data: dict, file_type: Optional[str]) -> Tuple[Dict[int, Any], bool, int]: ...

    def make_num_dict_rewrite(self, data: dict, file_type: Optional[str],
                              rewrite: bool) -> Tuple[Dict[int, Any], bool, int]: ...

    def get_key(self, num: int) -> Union[str, int]:
        """返回指定列序号对应的header key，如为None返回列序号
        :param num: 列序号
        :return: 表头值或列序号
        """
        ...

    def get_num(self, col: Union[int, str], is_header: bool = True) -> Optional[int]:
        """获取某列序号
        :param col: 列号、列名、表头值
        :param is_header: 当col为str时，是header的key还是列名
        :return: 列号int
        """
        ...


class ZeroHeader(Header):
    _OBJ: ZeroHeader = ...


class RowData(dict):
    header: Header = ...
    row: int = ...
    None_val: Optional[''] = ...

    def __init__(self, row: int, header: Header, None_val: Optional[''], seq: dict): ...

    def val(self, key: Union[int, str], is_header: bool = False, coord: bool = False) -> Any:
        """当前行获取指定列的值
        :param key: 为int时表示列序号，为str时表示列号或header key
        :param is_header: 为str时是header key还是列号
        :param coord: 为True时返回结果带坐标
        :return: coord为False时返回指定列的值，为Ture时返回(坐标, 值)
        """
        ...

    def col(self, key: str, num: bool = False) -> Union[str, int]:
        """获取指定表头项数据所在列
        :param key: 表头项
        :param num: 为True时返回列序号，否则返回列号
        :return: coord为False时返回指定列的值，为Ture时返回(坐标, 值)
        """
        ...


def align_csv(path: Union[str, Path], encoding: str = 'utf-8', delimiter: str = ',', quotechar: str = '"') -> None:
    """补全csv文件，使其每行列数一样多，用于pandas读取时避免出错
    :param path: 要处理的文件路径
    :param encoding: 文件编码
    :param delimiter: 分隔符
    :param quotechar: 引用符
    :return: None
    """
    ...


def get_usable_path(path: Union[str, Path], is_file: bool = True, parents: bool = True) -> Path:
    """检查文件或文件夹是否有重名，并返回可以使用的路径
    :param path: 文件或文件夹路径
    :param is_file: 目标是文件还是文件夹
    :param parents: 是否创建目标路径
    :return: 可用的路径，Path对象
    """
    ...


def make_valid_name(full_name: str) -> str:
    """获取有效的文件名
    :param full_name: 文件名
    :return: 可用的文件名
    """
    ...


def get_long(txt: str) -> int:
    """返回字符串中字符个数（一个汉字是2个字符）
    :param txt: 字符串
    :return: 字符个数
    """
    ...


def parse_coord(coord: Union[int, str, list, tuple, None],
                data_col: Optional[int]) -> Tuple[Optional[int], Optional[int]]:
    """添加数据，每次添加一行数据，可指定坐标、列号或行号
    coord只输入数字（行号）时，列号为self.data_col值，如 3；
    输入列号，或没有行号的坐标时，表示新增一行，列号为此时指定的，如'c'、',3'、(None, 3)、'None,3'；
    输入 'newline' 时，表示新增一行，列号为self.data_col值；
    输入行列坐标时，填写到该坐标，如'a3'、'3,1'、(3,1)、[3,1]；
    输入的行号可以是负数（列号不可以），代表从下往上数，-1是倒数第一行，如'a-3'、(-3, 3)
    :param coord: 坐标、列号、行号
    :param data_col: 列号，用于只传入行号的情况
    :return: 坐标tuple：(行, 列)坐标中的None表示新行或列
    """
    ...


def process_content_xlsx(content: Any) -> Union[None, int, str, float]:
    """处理单个单元格要写入的数据
    :param content: 未处理的数据内容
    :return: 处理后的数据
    """
    ...


def process_content_json(content: Any) -> Union[None, int, str, float]:
    """处理单个单元格要写入的数据
    :param content: 未处理的数据内容
    :return: 处理后的数据
    """
    ...


def process_content_str(content: Any) -> str:
    """处理单个单元格要写入的数据，以str格式输出
    :param content: 未处理的数据内容
    :return: 处理后的数据
    """
    ...


def ok_list_xlsx(data_list: Iterable) -> list:
    """处理列表中数据使其符合保存规范
    :param data_list: 数据列表
    :return: 处理后的列表
    """
    ...


def ok_list_str(data_list: Iterable) -> list:
    """处理列表中数据使其符合保存规范，所有数据都是str格式
    :param data_list: 数据列表
    :return: 处理后的列表
    """
    ...


def ok_list_db(data_list: Iterable) -> list:
    """处理列表中数据使其符合保存规范
    :param data_list: 数据列表
    :return: 处理后的列表
    """
    ...


def get_real_col(col: int, max_col: int):
    """获取返回真正写入文件的列序号
    :param col: 输入的列序号
    :param max_col: 最大列号
    :return: 真正的列序号
    """
    ...


def get_real_coord(coord: Union[tuple, list],
                   max_row: int,
                   max_col: Union[int, Worksheet]) -> Tuple[int, int]:
    """返回真正写入文件的坐标
    :param coord: 已初步格式化的坐标，如(1, 2)、(None, 3)、(-3, -2)
    :param max_row: 文件最大行
    :param max_col: 文件最大列
    :return: 真正写入文件的坐标，tuple格式
    """
    ...


def data_to_list_or_dict_simplify(recorder: BaseRecorder,
                                  data: Union[list, tuple, dict, None]) -> Union[list, dict]:
    """将传入的数据转换为列表或字典形式，不添加前后列数据
    :param recorder: BaseRecorder对象
    :param data: 要处理的数据
    :return: 转变成列表或字典形式的数据
    """
    ...


def data_to_list_or_dict(recorder: BaseRecorder, data: Iterable) -> Union[list, dict]:
    """将传入的一维数据转换为列表或字典形式，添加前后列数据
    :param recorder: BaseRecorder对象
    :param data: 要处理的数据
    :return: 转变成列表或字典形式的数据
    """
    ...


def get_csv(recorder: Recorder) -> Tuple[TextIOWrapper, bool]: ...


def get_wb(recorder: Recorder) -> tuple: ...


def get_ws(wb: Workbook, table, tables, new_file) -> Tuple[Worksheet, bool]: ...


def get_tables(path: Union[str, Path]) -> list: ...


def do_nothing(*args, **kwargs) -> None: ...


def get_key_cols(cols: Union[str, int, list, tuple, bool], header: Header, is_header: bool) -> List[int]:
    """获取作为关键字的列，可以是多列
    :param cols: 列号或列名，或它们组成的list或tuple
    :param header: Header格式
    :param is_header: cols中的str表示header还是列名
    :return: 列序号列表
    """
    ...


def is_sigal_data(data: Any) -> bool:
    """判断数据是否独立数据"""
    ...


def is_1D_data(data: Any) -> bool:
    """判断传入数据是否一维数据"""
    ...
