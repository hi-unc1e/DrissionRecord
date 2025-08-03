# -*- coding:utf-8 -*-
from threading import Lock
from typing import Literal, Optional, Any, Union

from openpyxl.cell import Cell
from openpyxl.styles import Alignment, Font, Border, Fill, Protection, Side, PatternFill, Color

LINES = Literal['dashDot', 'dashDotDot', 'dashed', 'dotted', 'double', 'hair', 'medium', 'mediumDashDot',
'mediumDashDotDot', 'mediumDashed', 'slantDashDot', 'thick', 'thin', None]


class CellStyle(object):
    font_args: tuple = ...
    border_args: tuple = ...
    alignment_args: tuple = ...
    protection_args: tuple = ...
    gradient_fill_args: tuple = ...
    pattern_fill_args: tuple = ...

    def __init__(self) -> None:
        """用于管理单元格样式的类"""
        self._font: Optional[CellFont] = ...
        self._border: Optional[CellBorder] = ...
        self._pattern_fill: Optional[CellPatternFill] = ...
        self._gradient_fill: Optional[CellGradientFill] = ...
        self._number_format: Optional[CellNumberFormat] = ...
        self._protection: Optional[CellProtection] = ...
        self._alignment: Optional[CellAlignment] = ...
        self._Font: Optional[Font] = ...
        self._Border: Optional[Border] = ...
        self._Alignment: Optional[Alignment] = ...
        self._Fill: Optional[Fill] = ...
        self._Protection: Optional[Protection] = ...
        self.height: Optional[float] = ...
        self.width: Optional[float] = ...

    @property
    def font(self) -> CellFont:
        """返回用于设置单元格字体的对象"""
        ...

    @property
    def border(self) -> CellBorder:
        """返回用于设置单元格边框的对象"""
        ...

    @property
    def alignment(self) -> CellAlignment:
        """返回用于设置单元格对齐选项的对象"""
        ...

    @property
    def pattern_fill(self) -> CellPatternFill:
        """返回用于设置单元格图案填充的对象"""
        ...

    @property
    def gradient_fill(self) -> CellGradientFill:
        """返回用于设置单元格渐变填充的对象"""
        ...

    @property
    def number_format(self) -> CellNumberFormat:
        """返回用于设置单元格数字格式的对象"""
        ...

    @property
    def protection(self) -> CellProtection:
        """返回用于设置单元格保护选项的对象"""
        ...

    def to_cell(self, cell: Cell, replace: bool = True) -> None:
        """把当前样式复制到目标单元格
        :param cell: 被设置样式的单元格对象
        :param replace: 是否直接替换目标单元格的样式，是的话效率较高，但不能保留未被设置的原有样式项
        :return: None
        """
        ...

    def set_size(self, width: float = None, height: float = None) -> CellStyle:
        """设置单元格宽和高，width和height都是None时清除宽高设置
        :param width: 单元格宽度，为None则不改变
        :param height: 单元格高度，为None则不改变
        :return: None
        """
        ...

    def set_bgColor(self, color: Union[None, str, tuple, Color]) -> CellStyle:
        """设置背景颜色
        :param color: 格式：'FFFFFF', '255,255,255', (255, 255, 255), Color对象均可，None表示恢复默认
        :return: 样式对象本身
        """
        ...

    def set_txtSize(self, size: float = None, bold: bool = None) -> CellStyle:
        """设置字体大小和是否加粗
        :param size: 大小，为None时不修改
        :param bold: 是否加粗，为None时不修改
        :return: 样式对象本身
        """
        ...

    def set_txtColor(self, color: Union[None, str, tuple, Color]) -> CellStyle:
        """设置文本颜色
        :param color: 字体颜色，格式：'FFFFFF', '255,255,255', (255, 255, 255), Color对象均可，None表示恢复默认
        :return: 样式对象本身
        """
        ...

    def set_bold(self, on_off: bool = True) -> CellStyle:
        """设置字体是否加粗
        :param on_off: bool表示有或无
        :return: 样式对象本身
        """
        ...

    def set_delLine(self, on_off: bool = True) -> CellStyle:
        """设置是否有删除线
        :param on_off: bool表示有或无
        :return: 样式对象本身
        """
        ...

    def set_underLine(self, on_off: bool = True) -> CellStyle:
        """设置是否有下划线线
        :param on_off: bool表示有或无
        :return: 样式对象本身
        """
        ...

    def set_center(self) -> CellStyle:
        """设置横向和竖向都居中
        :return: 样式对象本身
        """
        ...

    def set_border(self, on_off: bool = True) -> CellStyle:
        """设置边框
        :param on_off: bool表示开关
        :return: 样式对象本身
        """
        ...

    def _cover_to_cell(self, cell: Cell) -> None:
        """把当前样式复制到目标单元格，只覆盖有设置的项，没有设置的原有的项不变
        :param cell: 被设置样式的单元格对象
        :return: None
        """
        ...

    def _replace_to_cell(self, cell: Cell) -> None:
        """把当前样式复制到目标单元格，覆盖原有的设置
        :param cell: 被设置样式的单元格对象
        :return: None
        """
        ...


def _handle_args(args: tuple, src: Any, target: Any) -> dict: ...


class CellFont(object):
    _LINE_STYLES: tuple = ...
    _SCHEMES: tuple = ...
    _VERT_ALIGNS: tuple = ...

    def __init__(self):
        self.name: str = ...
        self.charset: int = ...
        self.size: float = ...
        self.bold: bool = ...
        self.italic: bool = ...
        self.strike: bool = ...
        self.outline: bool = ...
        self.shadow: bool = ...
        self.condense: bool = ...
        self.extend: bool = ...
        self.underline: Literal['single', 'double', 'singleAccounting', 'doubleAccounting'] = ...
        self.vertAlign: Literal['superscript', 'subscript', 'baseline'] = ...
        self.color: Union[Color, str] = ...
        self.scheme: Literal['major', 'minor'] = ...

    def set_name(self, name: Optional[str]) -> None:
        """设置字体
        :param name: 字体名称，None表示恢复默认
        :return: None
        """
        ...

    def set_charset(self, charset: Optional[int]) -> None:
        """设置编码
        :param charset: 字体编码，int格式，None表示恢复默认
        :return: None
        """
        ...

    def set_size(self, size: Optional[float]) -> None:
        """设置字体大小
        :param size: 字体大小，None表示恢复默认
        :return: None
        """
        ...

    def set_bold(self, on_off: Optional[bool]) -> None:
        """设置是否加粗
        :param on_off: bool表示开关，None表示恢复默认
        :return: None
        """
        ...

    def set_italic(self, on_off: Optional[bool]) -> None:
        """设置是否斜体
        :param on_off: bool表示开关，None表示恢复默认
        :return: None
        """
        ...

    def set_strike(self, on_off: Optional[bool]) -> None:
        """设置是否有删除线
        :param on_off: bool表示开关，None表示恢复默认
        :return: None
        """
        ...

    def set_outline(self, on_off: Optional[bool]) -> None:
        """设置outline
        :param on_off: bool表示开关，None表示恢复默认
        :return: None
        """
        ...

    def set_shadow(self, on_off: Optional[bool]) -> None:
        """设置是否有阴影
        :param on_off: bool表示开关，None表示恢复默认
        :return: None
        """
        ...

    def set_condense(self, on_off: Optional[bool]) -> None:
        """设置condense
        :param on_off: bool表示开关，None表示恢复默认
        :return: None
        """
        ...

    def set_extend(self, on_off: Optional[bool]) -> None:
        """设置extend
        :param on_off: bool表示开关，None表示恢复默认
        :return: None
        """
        ...

    def set_color(self, color: Union[None, str, tuple, Color]) -> None:
        """设置字体颜色
        :param color: 字体颜色，格式：'FFFFFF', '255,255,255', (255, 255, 255), Color对象均可，None表示恢复默认
        :return: None
        """
        ...

    def set_underline(self,
                      option: Literal['single', 'double', 'singleAccounting', 'doubleAccounting', None]) -> None:
        """设置下划线
        :param option: 下划线类型，可选 'single', 'double', 'singleAccounting', 'doubleAccounting'，None表示恢复默认
        :return: None
        """
        ...

    def set_vertAlign(self, option: Literal['superscript', 'subscript', 'baseline', None]) -> None:
        """设置上下标
        :param option: 可选 'superscript', 'subscript', 'baseline'，None表示恢复默认
        :return: None
        """
        ...

    def set_scheme(self, option: Literal['major', 'minor', None]) -> None:
        """设置scheme
        :param option: 可选 'major', 'minor'，None表示恢复默认
        :return: None
        """
        ...


class CellBorder(object):
    _LINE_STYLES: tuple = ...

    def __init__(self):
        self.start: Side = ...
        self.end: Side = ...
        self.left: Side = ...
        self.right: Side = ...
        self.top: Side = ...
        self.bottom: Side = ...
        self.diagonal: Side = ...
        self.vertical: Side = ...
        self.horizontal: Side = ...
        self.horizontal: Side = ...
        self.outline: bool = ...
        self.diagonalUp: bool = ...
        self.diagonalDown: bool = ...

    def set_start(self, style: LINES, color: Union[None, str, tuple, Color]) -> None:
        """设置start
        :param style: 线形，'dashDot','dashDotDot', 'dashed','dotted', 'double','hair', 'medium', 'mediumDashDot',
                      'mediumDashDotDot', 'mediumDashed', 'slantDashDot', 'thick', 'thin'，None表示恢复默认
        :param color: 边框颜色，格式：'FFFFFF', '255,255,255', (255, 255, 255), Color对象均可，None表示恢复默认
        :return: None
        """
        ...

    def set_end(self, style: LINES, color: Union[None, str, tuple, Color]) -> None:
        """设置end
        :param style: 线形，'dashDot','dashDotDot', 'dashed','dotted', 'double','hair', 'medium', 'mediumDashDot',
                      'mediumDashDotDot', 'mediumDashed', 'slantDashDot', 'thick', 'thin'，None表示恢复默认
        :param color: 边框颜色，格式：'FFFFFF', '255,255,255', (255, 255, 255), Color对象均可，None表示恢复默认
        :return: None
        """
        ...

    def set_left(self, style: LINES, color: Union[None, str, tuple, Color]) -> None:
        """设置左边框
        :param style: 线形，'dashDot','dashDotDot', 'dashed','dotted', 'double','hair', 'medium', 'mediumDashDot',
                      'mediumDashDotDot', 'mediumDashed', 'slantDashDot', 'thick', 'thin'，None表示恢复默认
        :param color: 边框颜色，格式：'FFFFFF', '255,255,255', (255, 255, 255), Color对象均可，None表示恢复默认
        :return: None
        """
        ...

    def set_right(self, style: LINES, color: Union[None, str, tuple, Color]) -> None:
        """设置右边框
        :param style: 线形，'dashDot','dashDotDot', 'dashed','dotted', 'double','hair', 'medium', 'mediumDashDot',
                      'mediumDashDotDot', 'mediumDashed', 'slantDashDot', 'thick', 'thin'，None表示恢复默认
        :param color: 边框颜色，格式：'FFFFFF', '255,255,255', (255, 255, 255), Color对象均可，None表示恢复默认
        :return: None
        """
        ...

    def set_top(self, style: LINES, color: Union[None, str, tuple, Color]) -> None:
        """设置上边框
        :param style: 线形，'dashDot','dashDotDot', 'dashed','dotted', 'double','hair', 'medium', 'mediumDashDot',
                      'mediumDashDotDot', 'mediumDashed', 'slantDashDot', 'thick', 'thin'，None表示恢复默认
        :param color: 边框颜色，格式：'FFFFFF', '255,255,255', (255, 255, 255), Color对象均可，None表示恢复默认
        :return: None
        """
        ...

    def set_bottom(self, style: LINES, color: Union[None, str, tuple, Color]) -> None:
        """设置下边框
        :param style: 线形，'dashDot','dashDotDot', 'dashed','dotted', 'double','hair', 'medium', 'mediumDashDot',
                      'mediumDashDotDot', 'mediumDashed', 'slantDashDot', 'thick', 'thin'，None表示恢复默认
        :param color: 边框颜色，格式：'FFFFFF', '255,255,255', (255, 255, 255), Color对象均可，None表示恢复默认
        :return: None
        """
        ...

    def set_quad(self, style: LINES, color: Union[None, str, tuple, Color]) -> None:
        """设置四个边框边框
        :param style: 线形，'dashDot','dashDotDot', 'dashed','dotted', 'double','hair', 'medium', 'mediumDashDot',
                      'mediumDashDotDot', 'mediumDashed', 'slantDashDot', 'thick', 'thin'，None表示恢复默认
        :param color: 边框颜色，格式：'FFFFFF', '255,255,255', (255, 255, 255), Color对象均可，None表示恢复默认
        :return: None
        """
        ...

    def set_diagonal(self, style: LINES, color: Union[None, str, tuple, Color]) -> None:
        """设置对角线
        :param style: 线形，'dashDot','dashDotDot', 'dashed','dotted', 'double','hair', 'medium', 'mediumDashDot',
                      'mediumDashDotDot', 'mediumDashed', 'slantDashDot', 'thick', 'thin'，None表示恢复默认
        :param color: 边框颜色，格式：'FFFFFF', '255,255,255', (255, 255, 255), Color对象均可，None表示恢复默认
        :return: None
        """
        ...

    def set_vertical(self, style: LINES, color: Union[None, str, tuple, Color]) -> None:
        """设置垂直中线
        :param style: 线形，'dashDot','dashDotDot', 'dashed','dotted', 'double','hair', 'medium', 'mediumDashDot',
                      'mediumDashDotDot', 'mediumDashed', 'slantDashDot', 'thick', 'thin'，None表示恢复默认
        :param color: 边框颜色，格式：'FFFFFF', '255,255,255', (255, 255, 255), Color对象均可，None表示恢复默认
        :return: None
        """
        ...

    def set_horizontal(self, style: LINES, color: Union[None, str, tuple, Color]) -> None:
        """设置水平中线
        :param style: 线形，'dashDot','dashDotDot', 'dashed','dotted', 'double','hair', 'medium', 'mediumDashDot',
                      'mediumDashDotDot', 'mediumDashed', 'slantDashDot', 'thick', 'thin'，None表示恢复默认
        :param color: 边框颜色，格式：'FFFFFF', '255,255,255', (255, 255, 255), Color对象均可，None表示恢复默认
        :return: None
        """
        ...

    def set_outline(self, on_off: bool) -> None:
        """
        :param on_off: bool表示开关
        :return: None
        """
        ...

    def set_diagonalDown(self, on_off: bool) -> None:
        """
        :param on_off: bool表示开关
        :return: None
        """
        ...

    def set_diagonalUp(self, on_off: bool) -> None:
        """
        :param on_off: bool表示开关
        :return: None
        """
        ...


class CellAlignment(object):
    _horizontal_alignments: tuple = ...
    _vertical_alignments: tuple = ...

    def __init__(self):
        self.horizontal = 'notSet'
        self.vertical = 'notSet'
        self.indent = 'notSet'
        self.relativeIndent = 'notSet'
        self.justifyLastLine = 'notSet'
        self.readingOrder = 'notSet'
        self.textRotation = 'notSet'
        self.wrapText = 'notSet'
        self.shrinkToFit = 'notSet'

    def set_horizontal(self,
                       horizontal: Literal['general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous',
                       'distributed', None]) -> None:
        """设置水平位置
        :param horizontal: 可选：'general', 'left', 'center', 'right', 'fill', 'justify', 'centerContinuous',
                                'distributed'，None表示恢复默认
        :return: None
        """
        ...

    def set_vertical(self, vertical: Literal['top', 'center', 'bottom', 'justify', 'distributed', None]) -> None:
        """设置垂直位置
        :param vertical: 可选：'top', 'center', 'bottom', 'justify', 'distributed'，None表示恢复默认
        :return: None
        """
        ...

    def set_indent(self, indent: int) -> None:
        """设置缩进
        :param indent: 缩进数值，0到255
        :return: None
        """
        ...

    def set_relativeIndent(self, indent: int) -> None:
        """设置相对缩进
        :param indent: 缩进数值，-255到255
        :return: None
        """
        ...

    def set_justifyLastLine(self, on_off: Optional[bool]) -> None:
        """设置justifyLastLine
        :param on_off: bool表示开或关，None表示恢复默认
        :return: None
        """
        ...

    def set_readingOrder(self, value: int) -> None:
        """设置readingOrder
        :param value: 不小于0的数字
        :return: None
        """
        ...

    def set_textRotation(self, value: int) -> None:
        """设置文本旋转角度
        :param value: 0-180或255
        :return: None
        """
        ...

    def set_wrapText(self, on_off: Optional[bool]) -> None:
        """设置wrapText
        :param on_off: bool表示开或关，None表示恢复默认
        :return: None
        """
        ...

    def set_shrinkToFit(self, on_off: Optional[bool]) -> None:
        """设置shrinkToFit
        :param on_off: bool表示开或关，None表示恢复默认
        :return: None
        """
        ...


class CellGradientFill(object):
    def __init__(self):
        self.type: str = ...
        self.degree: float = ...
        self.left: float = ...
        self.right: float = ...
        self.top: float = ...
        self.bottom: float = ...
        self.stop: Union[list, tuple] = ...

    def set_type(self, name: Literal['linear', 'path']) -> None:
        """设置类型
        :param name: 可选：'linear', 'path'
        :return: None
        """
        ...

    def set_degree(self, value: float) -> None:
        """设置程度
        :param value: 数值
        :return: None
        """
        ...

    def set_left(self, value: float) -> None:
        """设置left
        :param value: 数值
        :return: None
        """
        ...

    def set_right(self, value: float) -> None:
        """设置right
        :param value: 数值
        :return: None
        """
        ...

    def set_top(self, value: float) -> None:
        """设置top
        :param value: 数值
        :return: None
        """
        ...

    def set_bottom(self, value: float) -> None:
        """设置bottom
        :param value: 数值
        :return: None
        """
        ...

    def set_stop(self, values: Union[list, tuple]) -> None:
        """设置stop
        :param values: 数值
        :return: None
        """
        ...


class CellPatternFill(object):
    _FILES: tuple = ...

    def __init__(self):
        self.patternType: str = ...
        self.fgColor: Union[str, Color] = ...
        self.bgColor: Union[str, Color] = ...

    def set_patternType(self, name: Literal[
        'none', 'solid', 'darkDown', 'darkGray', 'darkGrid', 'darkHorizontal', 'darkTrellis', 'darkUp',
        'darkVertical', 'gray0625', 'gray125', 'lightDown', 'lightGray', 'lightGrid', 'lightHorizontal',
        'lightTrellis', 'lightUp', 'lightVertical', 'mediumGray', None]) -> None:
        """设置类型
        :param name: 可选：'none', 'solid', 'darkDown', 'darkGray', 'darkGrid', 'darkHorizontal', 'darkTrellis',
                          'darkUp', 'darkVertical', 'gray0625', 'gray125', 'lightDown', 'lightGray', 'lightGrid',
                          'lightHorizontal', 'lightTrellis', 'lightUp', 'lightVertical', 'mediumGray'，None为恢复默认
        :return: None
        """
        ...

    def set_fgColor(self, color: Union[None, str, tuple, Color]) -> None:
        """设置前景色
        :param color: 颜色，格式：'FFFFFF', '255,255,255', (255, 255, 255), Color对象均可，None表示恢复默认
        :return: None
        """
        ...

    def set_bgColor(self, color: Union[None, str, tuple, Color]) -> None:
        """设置背景色
        :param color: 颜色，格式：'FFFFFF', '255,255,255', (255, 255, 255), Color对象均可，None表示恢复默认
        :return: None
        """
        ...


class CellNumberFormat(object):
    def __init__(self):
        self.format: str = 'notSet'

    def set_format(self, string: Optional[str]) -> None:
        """设置数字格式
        :param string: 格式字符串，为None时恢复默认，格式很多具体在`openpyxl.numbers`查看
        :return: None
        """
        ...


class CellProtection(object):
    def __init__(self):
        self.hidden: bool = ...
        self.locked: bool = ...

    def set_hidden(self, on_off: bool) -> None:
        """设置是否隐藏
        :param on_off: bool表示开关
        :return: None
        """
        ...

    def set_locked(self, on_off: bool) -> None:
        """设置是否锁定
        :param on_off: bool表示开关
        :return: None
        """
        ...


class CellStyleCopier(object):
    def __init__(self, from_cell: Cell):
        self._style = ...
        self._font: Font = ...
        self._border: Border = ...
        self._fill: Fill = ...
        self._number_format = ...
        self._protection: Protection = ...
        self._alignment: Alignment = ...

    def to_cell(self, cell: Cell) -> None:
        """把当前样式复制到目标单元格
        :param cell: 被设置样式的单元格对象
        :return: None
        """
        ...


def get_color_code(color: Union[str, tuple, Color]) -> str:
    """将颜色拼音转为代码
    :param color: 颜色名称或代码字符串
    :return: 颜色代码
    """
    ...


class NoneStyle(object):
    _instance_lock: Lock = ...

    def __init__(self):
        self._font: Font = ...
        self._border: Border = ...
        self._alignment: Alignment = ...
        self._fill: PatternFill = ...
        self._number_format: str = ...
        self._protection: Protection = ...

    def __new__(cls, *args, **kwargs): ...

    def to_cell(self, cell: Cell, replace: bool = True) -> None:
        """把当前样式复制到目标单元格
        :param cell: 被设置样式的单元格对象
        :param replace: 用于对齐
        :return: None
        """
        ...
