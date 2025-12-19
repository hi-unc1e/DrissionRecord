# -*- coding:utf-8 -*-
from io import TextIOWrapper
from pathlib import Path
from typing import Union, Tuple, Any, Optional, List, Dict, Iterable, Literal

from openpyxl.workbook import Workbook
from openpyxl.worksheet.worksheet import Worksheet

from .base import BaseRecorder
from .cell_style import CellStyleCopier
from .recorder import Recorder

REWRITE_METHOD = Literal['make_num_dict_rewrite', 'make_num_dict']


def line2ws(ws: Worksheet, header: Header, row: int, col: int, data: Union[dict, list], rewrite_method: REWRITE_METHOD,
            rewrite: bool) -> bool:
    """把一行数据写入数据表，不设置样式
    :param ws: Worksheet对象
    :param header: Header对象
    :param row: 行号
    :param col: 列序号
    :param data: 行数据
    :param rewrite_method: 'make_num_dict_rewrite'或'make_num_dict'
    :param rewrite: 是否重写表头
    :return: 是否重写表头
    """
    ...


def line2ws_follow(ws: Worksheet, header: Header, row: int, col: int, data: Union[dict, list],
                   rewrite_method: REWRITE_METHOD, rewrite: bool, styles: Dict[int, CellStyleCopier],
                   height: Optional[float], new_row: bool) -> bool:
    """把一行数据写入数据表，并设置样式
    :param ws: Worksheet对象
    :param header: Header对象
    :param row: 行号
    :param col: 列序号
    :param data: 行数据
    :param rewrite_method: 'make_num_dict_rewrite'或'make_num_dict'
    :param rewrite: 是否重写表头 
    :param styles: 样式对象了列表
    :param height: 行高，仅新行时有效
    :param new_row: 是否新行
    :return: 是否重写表头
    """
    ...


def data2ws(recorder: Recorder, ws: Worksheet, data: dict, coord: Tuple[int, int],
            header: Header, rewrite: bool, rewrite_method: REWRITE_METHOD) -> bool:
    """数据写入数据表
    :param recorder: Recorder对象
    :param ws: Worksheet对象
    :param data: 标准数据 {'type': 'data', 'data': [(1, 2, 3, 4)], 'coord': (0, 1)}
    :param coord: 要写入的坐标
    :param header: Header对象
    :param rewrite: 是否重写表头
    :param rewrite_method: 'make_num_dict_rewrite'或'make_num_dict'
    :return: 是否重写表头
    """
    ...


def data2ws_follow(recorder: Recorder, ws: Worksheet, data: dict, coord: Tuple[int, int],
                   header: Header, rewrite: bool, rewrite_method: REWRITE_METHOD) -> None:
    """数据写入数据表，跟随上一行样式
    :param recorder: Recorder对象
    :param ws: Worksheet对象
    :param data: 标准数据 {'type': 'data', 'data': [(1, 2, 3, 4)], 'coord': (0, 1)}
    :param coord: 要写入的坐标
    :param header: Header对象
    :param rewrite: 是否重写表头
    :param rewrite_method: 'make_num_dict_rewrite'或'make_num_dict'
    :return: 是否重写表头
    """
    ...


def data2ws_style(recorder: Recorder, ws: Worksheet, data: dict, coord: Tuple[int, int],
                  header: Header, rewrite: bool, rewrite_method: REWRITE_METHOD) -> None:
    """数据写入数据表，并设置指定样式
    :param recorder: Recorder对象
    :param ws: Worksheet对象
    :param data: 标准数据 {'type': 'data', 'data': [(1, 2, 3, 4)], 'coord': (0, 1)}
    :param coord: 要写入的坐标
    :param header: Header对象
    :param rewrite: 是否重写表头
    :param rewrite_method: 'make_num_dict_rewrite'或'make_num_dict'
    :return: 是否重写表头
    """
    ...


def styles2new_row(ws: Worksheet, styles: Iterable, height: float, row: int) -> None:
    """
    :param ws: Worksheet对象
    :param styles: Style对象租车的列表
    :param height: 行高
    :param row: 行号
    :return: None
    """
    ...


def styles2ws(**kwargs) -> None:
    """把样式写入数据表"""
    ...


def link2ws(**kwargs) -> None:
    """把擦后入到单元格"""
    ...


def img2ws(**kwargs) -> None:
    """把图片到单元格"""
    ...


def width2ws(**kwargs) -> None:
    """把列宽设置到数据表"""
    ...


def height2ws(**kwargs) -> None:
    """把行高设置到数据表"""
    ...


def remove_end_Nones(in_list: list) -> list:
    """去除列表后面所有None
    :param in_list: 要处理的list
    :return: 处理后的列表
    """
    ...


class BaseHeader(object):
    _NUM_KEY: dict = ...
    _KEY_NUM: dict = ...
    _CONTENT_FUNCS: dict = ...

    @property
    def key_num(self) -> Dict[str, int]:
        """{str: int}格式的表头数据"""
        ...

    @property
    def num_key(self) -> Dict[int, str]:
        """{int: str}格式的表头数据"""
        ...

    def __iter__(self): ...


class Header(BaseHeader):

    def __init__(self, header: Iterable = None): ...

    def __getitem__(self, item: Union[int, str]): ...

    def __len__(self) -> int: ...

    def values(self):
        """"""
        ...

    def items(self):
        """"""
        ...

    def make_row_data(self, row: int, row_values: dict, None_val: Optional[''] = None) -> RowData:
        """生成RowData对象
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

    def make_num_dict(self, *keys) -> Tuple[Dict[int, Any], bool, int]:
        """生成{int: val}的行数据，不考虑是否重写表头
        :return: (处理后的行数据, 是否重写表头, 表头长度)
        """
        ...

    def make_num_dict_rewrite(self, *keys) -> Tuple[Dict[int, Any], bool, int]:
        """生成{int: val}的行数据，虑是否重写表头
        :return: (处理后的行数据, 是否重写表头, 表头长度)
        """
        ...

    def get_key(self, num: int) -> Union[str, int]:
        """返回指定列序号对应的表头值，如该列没有值，返回列序号
        :param num: 列序号
        :return: 表头值或列序号
        """
        ...

    def get_col(self, header_or_num) -> Optional[str]:
        """返回指定列序号或表头值对应的列号，无指定表头值时返回None
        :param header_or_num: 表头值或列序号
        :return: 列号'A'
        """
        ...

    def get_num(self, header_or_num: Union[int, str]) -> Optional[int]:
        """返回指定列序号或表头值对应的列序号，找不到表头值时返回None
        :param header_or_num: 列号、表头值
        :return: 列号int
        """
        ...

    def _get_num(self, header_or_num: Union[int, str]) -> int:
        """内部使用，返回指定列序号或表头值对应的列序号，找不到表头值时返回表头长度加1
        :param header_or_num: 列号、表头值
        :return: 列号int
        """
        ...

    def _num2num(self, num: int) -> int:
        """处理负数列序号，返回真实列序号，超出范围返回None，为0返回新列
        :param num: 列序号
        :return: 真实列号
        """
        ...


class ZeroHeader(Header):
    _OBJ: ZeroHeader = ...

    def _get_num(self, header_or_num: Union[int, str]) -> int:
        """返回指定列序号或表头值对应的列序号，找不到表头值时返回1
        :param header_or_num: 列号、表头值
        :return: 列号int
        """
        ...


class RowData(dict):
    header: Header = ...
    row: int = ...
    _None_val: Optional[''] = ...

    def __init__(self, row: int, header: Header, None_val: Optional[''], seq: dict): ...

    def col(self, key_or_num: Union[int, str], as_num: bool = True) -> Union[int, str]:
        """返回数据中指定列的列号或列序号
        :param key_or_num: 为int时表示列序号，为str时表示表头值
        :param as_num: 列以列号还是列序号形式返回
        :return: 返回列（'A'或1）
        """
        ...

    def coord(self, key_or_num: Union[int, str], col_num: bool = False) -> Tuple[int, Union[str, int]]:
        """返回数据中指定列的坐标
        :param key_or_num: 为int时表示列序号，为str时表示表头值
        :param col_num: 列以列号还是列序号形式返回
        :return: 返回(行号, 列号)
        """
        ...


class RowText(str):
    row: Optional[int] = ...


def Col(key: str) -> int:
    """输入列号，输出列序号
    :param key: 列号'A'
    :return: 第几列
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
                data_col: int = 1) -> Tuple[Optional[int], Optional[int]]:
    """处理坐标格式
    :param coord: 'A3'格式坐标、(3, 1)或(3, '列名')格式坐标、行号
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


def process_nothing(content: Any) -> Any:
    """不处理直接返回数据"""
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


def get_real_row(row: int, max_row: int) -> int:
    """获取返回真正写入文件的行号
    :param row: 输入的行号
    :param max_row: 最大行号
    :return: 真正的行号
    """
    ...


def get_real_coord(coord: tuple, max_row: int,
                   header: Header) -> Tuple[int, int]:
    """返回真正写入文件的坐标
    :param coord: 已初步格式化的坐标，如(1, 2)、(0, 3)、(-3, -2)
    :param max_row: 文件最大行
    :param header: Header对象
    :return: 真正写入文件的坐标，tuple格式
    """
    ...


def get_ws_real_coord(coord: tuple, ws: Worksheet, header: Header) -> Tuple[int, int]:
    """返回真正写入xlsx文件的坐标
    :param coord: 已初步格式化的坐标，如(1, 2)、(0, 3)、(-3, -2)
    :param ws: Worksheet对象
    :param header: Header对象
    :return: 真正写入文件的坐标，tuple格式
    """
    ...


def make_final_data_simplify(recorder: BaseRecorder,
                             data: Union[list, tuple, dict, None]) -> Union[list, dict]:
    """将传入的数据转换为列表或字典形式，不添加前后列数据
    :param recorder: BaseRecorder对象
    :param data: 要处理的数据
    :return: 转变成列表或字典形式的数据
    """
    ...


def make_final_data(recorder: BaseRecorder, data: Iterable) -> Union[list, dict]:
    """将传入的一维数据转换为列表或字典形式，添加前后列数据
    :param recorder: BaseRecorder对象
    :param data: 要处理的数据
    :return: 转变成列表或字典形式的数据
    """
    ...


def get_csv(recorder: Recorder) -> Tuple[TextIOWrapper, bool]:
    """获取文件读写对象
    :param recorder: Recorder对象
    :return: (文件读写对象, 是否新文件)
    """
    ...


def get_wb(recorder: Recorder) -> Tuple[Workbook, bool]:
    """获取Workbook对象
    :param recorder: Recorder对象
    :return: (Workbook对象, 是否新文件)
    """
    ...


def get_ws(wb: Workbook, table: Optional[str], tables: List[str], new_file: bool) -> Tuple[Worksheet, bool]:
    """获取Worksheet对象
    :param wb: Workbook对象
    :param table: 表名，None代表活动表格
    :param tables: 工作簿所有表名组成的列表
    :param new_file: 是否新文件
    :return: (Worksheet对象, 是否新文件)
    """
    ...


def get_tables(path: Union[str, Path]) -> List[str]:
    """获取所有数据表名称
    :param path: 文件路径
    :return: 表名组成的列表
    """
    ...


def do_nothing(*args, **kwargs) -> None:
    """什么都不干"""
    ...


def get_key_cols(cols: Union[str, int, list, tuple, bool], header: Header) -> List[int]:
    """获取作为关键字的列，可以是多列
    :param cols: 列号或列名，或它们组成的list或tuple
    :param header: Header格式
    :return: 列序号列表
    """
    ...


def is_single_data(data: Any) -> bool:
    """判断数据是否独立数据"""
    ...


def is_1D_data(data: Any) -> bool:
    """判断传入数据是否一维数据"""
    ...
