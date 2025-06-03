# -*- coding:utf-8 -*-
from csv import reader as csv_reader, writer as csv_writer
from pathlib import Path
from time import sleep

from openpyxl.reader.excel import load_workbook

from .base import BaseRecorder
from .setter import RecorderSetter, set_csv_header
from .cell_style import CellStyleCopier, CellStyle, NoneStyle
from .tools import (ok_list_xlsx, ok_list_str, process_content_xlsx, process_content_json, process_content_str,
                    fix_openpyxl_bug, get_key_cols, Header, get_wb, get_ws, get_csv, parse_coord,
                    data_to_list_or_dict, get_usable_coord, get_usable_coord_int, do_nothing, )


class Recorder(BaseRecorder):
    _SUPPORTS = ('csv', 'xlsx', 'txt', 'jsonl', 'json')

    def __init__(self, path=None, cache_size=1000):
        """用于缓存并记录数据，可在达到一定数量时自动记录，以降低文件读写次数，减少开销
        :param path: 保存的文件路径
        :param cache_size: 每接收多少条记录写入文件，0为不自动写入
        """
        self._header = {None: None}
        self._methods = {'xlsx': self._to_xlsx_fast,
                         'csv': self._to_csv_fast,
                         'txt': self._to_txt,
                         'jsonl': self._to_jsonl,
                         'json': self._to_json,
                         'add_data': self._add_data_fast,
                         'set_img': do_nothing,
                         'set_link': do_nothing,
                         'set_style': do_nothing,
                         'set_row_height': do_nothing,
                         'set_col_width': do_nothing,
                         }
        super().__init__(path=path, cache_size=cache_size)
        self._style_data = {}
        self._delimiter = ','  # csv文件分隔符
        self._quote_char = '"'  # csv文件引用符
        self._follow_styles = False
        self._row_height = None
        self._styles = None
        self._header_row = 1
        self._fast = True
        self._link_style = None
        self.data_col = 1

    def _set_methods(self, file_type):
        method = f'_to_{file_type}_fast' if hasattr(self, f'_to_{file_type}_fast') else f'_to_{file_type}'
        self._methods[file_type] = getattr(self, method)
        self._methods['add_data'] = self._add_data_fast
        if file_type == 'xlsx':
            self._methods['set_img'] = self._set_img
            self._methods['set_link'] = self._set_link
            self._methods['set_style'] = self._set_style
            self._methods['set_row_height'] = self._set_row_height
            self._methods['set_col_width'] = self._set_col_width
        else:
            self._methods['set_img'] = do_nothing
            self._methods['set_link'] = do_nothing
            self._methods['set_style'] = do_nothing
            self._methods['set_row_height'] = do_nothing
            self._methods['set_col_width'] = do_nothing

    @property
    def set(self):
        """返回用于设置属性的对象"""
        if self._setter is None:
            self._setter = RecorderSetter(self)
        return self._setter

    @property
    def delimiter(self):
        """返回csv文件分隔符"""
        return self._delimiter

    @property
    def quote_char(self):
        """返回csv文件引用符"""
        return self._quote_char

    @property
    def header(self):
        """返回表头，只支持csv和xlsx格式"""
        if self.type not in ('csv', 'xlsx'):
            raise TypeError('header属性只支持csv和xlsx类型文件。')
        return get_header(self)

    def add_data(self, data, coord=None, table=None):
        """添加数据，可一次添加多条数据
        :param data: 插入的数据，任意格式
        :param coord: 要添加数据的坐标，可输入行号、列号或行列坐标，当格式不是xlsx或csv时无效，eg.'a3'、1、[3, 1]、'c'
        :param table: 要写入的数据表，仅支持xlsx格式。为None表示用set.table()方法设置的值，为True表示活动的表格
        :return: None
        """
        while self._pause_add:  # 等待其它线程写入结束
            sleep(.02)
        if self._fast and coord and self._type in ('csv', 'xlsx'):
            self._to_slow_mode()
        self._methods['add_data'](data=data, coord=coord, table=table)

    def _add_data_fast(self, **args):
        da = args['data']
        if not isinstance(da, (list, tuple, dict)):
            da = (da,)

        if not da:
            da = ([],)
            self._data_count += 1

        # 一维数组
        elif isinstance(da, dict) or (isinstance(da, (list, tuple)) and not isinstance(da[0], (list, tuple, dict))):
            da = [self._handle_data_method(self, da)]
            self._data_count += 1

        else:  # 二维数组
            da = [self._handle_data_method(self, d) for d in da]
            self._data_count += len(da)

        if self._type == 'xlsx':
            table = args['table']
            if table is None:
                table = self._table
            elif table is True:
                table = None
            self._data.setdefault(table, []).extend(da)

        elif self._type:
            self._data.extend(da)

        else:
            raise RuntimeError('请设置文件路径。')

        if 0 < self.cache_size <= self._data_count:
            self.record()

    def _add_data_slow(self, **args):
        while self._pause_add:  # 等待其它线程写入结束
            sleep(.02)

        da = args['data']
        coord = args['coord']
        table = args['table']
        if not isinstance(da, (list, tuple)):
            da = (da,)

        to = self._data
        if coord in ('cover_style', 'replace_style', 'set_width', 'set_height'):
            to = self._style_data
            self._data_count += 1

        elif coord not in ('set_link', 'set_img'):
            coord = parse_coord(coord, self.data_col)
            if not da:
                da = ([],)
                self._data_count += 1
            # 一维数组
            elif isinstance(da, dict) or (isinstance(da, (list, tuple)) and not isinstance(da[0], (list, tuple, dict))):
                da = (data_to_list_or_dict(self, da),)
                self._data_count += 1
            else:  # 二维数组
                da = [self._handle_data_method(self, d) for d in da]
                self._data_count += len(da)

        else:
            self._data_count += 1

        if self._type == 'xlsx':
            if table is None:
                table = self._table
            elif isinstance(table, bool):
                table = None
            to.setdefault(table, []).append((coord, da))

        elif self._type:
            to.append((coord, da))

        else:
            raise RuntimeError('请设置文件路径。')

        if 0 < self.cache_size <= self._data_count:
            self.record()

    def set_link(self, coord, link, content=None, table=None):
        """为单元格设置超链接
        :param coord: 单元格坐标
        :param link: 超链接，为None时删除链接
        :param content: 单元格内容
        :param table: 数据表名，仅支持xlsx格式。为None表示用set.table()方法设置的值，为bool表示活动的表格
        :return: None
        """
        self._methods['set_link'](coord, link, content, table)

    def set_img(self, coord, img_path, width=None, height=None, table=None):
        """
        :param coord: 单元格坐标
        :param img_path: 图片路径
        :param width: 图片宽
        :param height: 图片高
        :param table: 数据表名，仅支持xlsx格式。为None表示用set.table()方法设置的值，为bool表示活动的表格
        :return: None
        """
        self._methods['set_img'](coord, img_path, width, height, table)

    def set_style(self, coord, style, replace=True, table=None):
        """为单元格设置样式，可批量设置范围内的单元格
        :param coord: 单元格坐标，输入数字可设置整行，输入列名字符串可设置整列，输入'A1:C5'、'a:d'、'1:5'格式可设置指定范围
        :param style: CellStyle对象，为None则清除单元格样式
        :param replace: 是否直接替换已有样式，运行效率较高，但不能单独修改某个属性
        :param table: 数据表名，仅支持xlsx格式。为None表示用set.table()方法设置的值，为bool表示活动的表格
        :return: None
        """
        self._methods['set_style'](coord, style, replace, table)

    def set_row_height(self, row, height, table=None):
        """设置行高，可设置连续多行
        :param row: 行号，可传入范围，如'1:4'
        :param height: 行高
        :param table: 数据表名，仅支持xlsx格式。为None表示用set.table()方法设置的值，为bool表示活动的表格
        :return: None
        """
        self._methods['set_row_height'](row, height, table)

    def set_col_width(self, col, width, table=None):
        """设置列宽，可设置连续多列
        :param col: 列号，数字或字母，可传入范围，如'1:4'、'a:d'
        :param width: 列宽
        :param table: 数据表名，仅支持xlsx格式。为None表示用set.table()方法设置的值，为bool表示活动的表格
        :return: None
        """
        self._methods['set_col_width'](col, width, table)

    def _set_link(self, coord, link, content=None, table=None):
        """为单元格设置超链接
        :param coord: 单元格坐标
        :param link: 超链接，为None时删除链接
        :param content: 单元格内容
        :param table: 数据表名，仅支持xlsx格式。为None表示用set.table()方法设置的值，为bool表示活动的表格
        :return: None
        """
        self.add_data((coord, link, content), 'set_link', table)

    def _set_img(self, coord, img_path, width=None, height=None, table=None):
        """
        :param coord: 单元格坐标
        :param img_path: 图片路径
        :param width: 图片宽
        :param height: 图片高
        :param table: 数据表名，仅支持xlsx格式。为None表示用set.table()方法设置的值，为bool表示活动的表格
        :return: None
        """
        self.add_data((coord, str(img_path), width, height), 'set_img', table)

    def _set_style(self, coord, style, replace=True, table=None):
        """为单元格设置样式，可批量设置范围内的单元格
        :param coord: 单元格坐标，输入数字可设置整行，输入列名字符串可设置整列，输入'A1:C5'、'a:d'、'1:5'格式可设置指定范围
        :param style: CellStyle对象，为None则清除单元格样式
        :param replace: 是否直接替换已有样式，运行效率较高，但不能单独修改某个属性
        :param table: 数据表名，仅支持xlsx格式。为None表示用set.table()方法设置的值，为bool表示活动的表格
        :return: None
        """
        s = 'replace_style' if replace else 'cover_style'
        self.add_data((coord, style), s, table)

    def _set_row_height(self, row, height, table=None):
        """设置行高，可设置连续多行
        :param row: 行号，可传入范围，如'1:4'
        :param height: 行高
        :param table: 数据表名，仅支持xlsx格式。为None表示用set.table()方法设置的值，为bool表示活动的表格
        :return: None
        """
        if not row:
            raise ValueError('row不能为0或None。')
        self.add_data((row, height), 'set_height', table)

    def _set_col_width(self, col, width, table=None):
        """设置列宽，可设置连续多列
        :param col: 列号，数字或字母，可传入范围，如'1:4'、'a:d'
        :param width: 列宽
        :param table: 数据表名，仅支持xlsx格式。为None表示用set.table()方法设置的值，为bool表示活动的表格
        :return: None
        """
        if not col:
            raise ValueError('col不能为0或None。')
        self.add_data((col, width), 'set_width', table)

    def rows(self, key_cols=True, sign_col=True, is_header=False,
             signs=None, deny_sign=False, count=None, begin_row=None):
        """返回符合条件的行数据，可指定只要某些列
        :param key_cols: 作为关键字的列，可以是多列，为True获取所有列
        :param sign_col: 用于筛选数据的列，为True获取所有行
        :param is_header: key_cols和sign_col是str时，表示header值还是列名
        :param signs: 按这个值判断是否已填数据，可用list, tuple, set设置多个
        :param deny_sign: 是否反向匹配sign，即筛选指不是sign的行
        :param count: 获取多少条数据，为None获取所有
        :param begin_row: 数据开始的行，None表示header_row后面一行
        :return: RowData对象
        """
        if not self._path or not Path(self._path).exists():
            raise RuntimeError('未指定文件路径或文件不存在。')
        if self.type == 'xlsx':
            wb = load_workbook(self.path, data_only=True, read_only=True)
            if self.table and self.table not in [i.title for i in wb.worksheets]:
                raise RuntimeError(f'xlsx文件未包含指定工作表：{self.table}')
            ws = wb[self.table] if self.table else wb.active
            if ws.max_column is None:  # 遇到过read_only时无法获取列数的文件
                wb.close()
                wb = load_workbook(self.path, data_only=True)
                ws = wb[self.table] if self.table else wb.active

            method = get_xlsx_rows

        elif self.type == 'csv':
            ws = None
            method = get_csv_rows
        else:
            raise RuntimeError('rows()方法只支持xlsx和csv格式。')

        header = get_header(self, ws)
        if not begin_row:
            begin_row = self._header_row + 1

        if sign_col is not True:
            sign_col = header.get_num(sign_col, is_header=is_header) or 1
        if not isinstance(signs, (list, tuple, set)):
            signs = (signs,)
        key_cols = get_key_cols(key_cols, header, is_header)

        return method(self, header=header, key_cols=key_cols, begin_row=begin_row, sign_col=sign_col,
                      sign=signs, deny_sign=deny_sign, count=count, ws=ws)

    def clear(self):
        super().clear()
        self._style_data.clear()

    def _record(self):
        """记录数据"""
        self._methods[self.type]()
        self._style_data = {}
        if not self._fast:
            self._to_fast_mode()

    def _to_fast_mode(self):
        self._methods['xlsx'] = self._to_xlsx_fast
        self._methods['csv'] = self._to_csv_fast
        self._methods['add_data'] = self._add_data_fast
        self._fast = True

    def _to_slow_mode(self):
        self._methods['xlsx'] = self._to_xlsx_slow
        self._methods['csv'] = self._to_csv_slow
        self._methods['add_data'] = self._add_data_slow
        self._fast = False

    def _to_xlsx_fast(self):
        """记录数据到xlsx文件"""
        wb, new_file = get_wb(self)
        tables = wb.sheetnames
        for table, data in self._data.items():
            _row_styles = None
            _row_height = None
            ws, new_sheet = get_ws(wb, table, tables, new_file)
            first_wrote = False
            if table is None and ws.title not in self._header:
                self._header[ws.title] = self._header[None]

            # ==============处理表头和样式==============
            if new_sheet:
                first_wrote = new_sheet_fast(self, ws, data, first_wrote)
            elif self._header.get(ws.title, None) is None:
                self._header[ws.title] = Header([c.value for c in ws[self._header_row]])

            header = self._header[ws.title]
            begin_row = None  # 开始写入数据的行
            if self._follow_styles:
                begin_row = ws.max_row
                _row_styles = [CellStyleCopier(i) for i in ws[begin_row]]
                _row_height = ws.row_dimensions[begin_row].height
                begin_row += 1
            elif self._styles or self._row_height:
                begin_row = ws.max_row + 1

            if new_file:
                wb, ws = fix_openpyxl_bug(self, wb, ws, ws.title)
                new_file = False

            # ==============开始写入数据==============
            if first_wrote:
                data = data[1:]

            rewrite_header = False
            for i in data:
                i, rewrite_header = header.make_insert_list(i, self._auto_new_header, rewrite_header, 'xlsx')
                ws.append(i)

            if rewrite_header:
                for c in range(1, ws.max_column + 1):
                    ws.cell(self._header_row, c, value=header[c])

            if self._follow_styles:
                for r in range(begin_row, ws.max_row + 1):
                    set_style(_row_height, _row_styles, ws, r)

            elif self._styles or self._row_height:
                if isinstance(self._styles, dict):
                    styles = header.make_num_dict(self._styles, None)
                    styles = [styles.get(c, None) for c in range(1, ws.max_column + 1)]

                elif isinstance(self._styles, CellStyle):
                    styles = [self._styles] * ws.max_column
                else:
                    styles = self._styles

                for r in range(begin_row, ws.max_row + 1):
                    set_style(self._row_height, styles, ws, r)

        wb.save(self.path)
        wb.close()

    def _to_xlsx_slow(self):
        """填写数据到xlsx文件"""
        wb, new_file = get_wb(self)
        tables = wb.sheetnames
        for table in {}.fromkeys(list(self._data.keys()) + list(self._style_data.keys())):
            ws, new_sheet = get_ws(wb, table, tables, new_file)
            first_data_wrote = False
            first_style_wrote = False
            if table is None and ws.title not in self._header:
                self._header[ws.title] = self._header[None]

            if new_sheet:
                first_data_wrote, first_style_wrote = new_sheet_slow(self, ws, self._data.get(table, []),
                                                                     self._style_data.get(table, []),
                                                                     first_data_wrote, first_style_wrote)
            elif self._header.get(ws.title, None) is None:
                self._header[ws.title] = Header([c.value for c in ws[self._header_row]])

            header = self._header[ws.title]
            if new_file:
                wb, ws = fix_openpyxl_bug(self, wb, ws, ws.title)
                new_file = False

            if self._data.get(table, None):
                method = data2ws_has_style if self._follow_styles else data2ws_no_style
                data = self._data[table][1:] if first_data_wrote else self._data[table]
                for cur_data in data:
                    if cur_data[0] == 'set_link':
                        set_link_to_ws(ws, cur_data[1], False, self)
                    elif cur_data[0] == 'set_img':
                        set_img_to_ws(ws, cur_data[1], False, self)
                    else:
                        max_row = ws.max_row
                        row, col = get_usable_coord(cur_data[0], max_row, ws)
                        not_new = cur_data[0][0]  # 是否添加到新行
                        cur_data = cur_data[1] if isinstance(cur_data[1][0], (list, tuple, dict)) else (cur_data[1],)
                        method(ws, header, row, col, cur_data, not_new, max_row)

            if self._style_data.get(table, None):
                style = self._style_data[table][1:] if first_style_wrote else self._style_data[table]
                for cur_data in style:
                    set_style_to_ws(ws, cur_data, False, self, header)

        wb.save(self.path)
        wb.close()

    def _to_csv_fast(self):
        """记录数据到csv文件"""
        file, new_csv = get_csv(self)
        writer = csv_writer(file, delimiter=self.delimiter, quotechar=self.quote_char)
        get_and_set_csv_header(self, new_csv, file, writer)

        rewrite_header = False
        for i in self._data:
            i, rewrite_header = self._header[None].make_insert_list(i, self._auto_new_header, rewrite_header, 'csv')
            writer.writerow(i)
        file.close()

        if rewrite_header:
            set_csv_header(self, self._header[None], self._header_row)

    def _to_csv_slow(self):
        """填写数据到csv文件"""
        file, new_csv = get_csv(self)
        writer = csv_writer(file, delimiter=self.delimiter, quotechar=self.quote_char)
        get_and_set_csv_header(self, new_csv, file, writer)
        file.seek(0)
        reader = csv_reader(file, delimiter=self.delimiter, quotechar=self.quote_char)
        lines = list(reader)
        lines_count = len(lines)
        header = self._header[None]

        rewrite_header = False
        for i in self._data:
            coord, cur_data = i
            row, col = get_usable_coord_int(coord, lines_count, len(lines[0]) if lines_count else 1)
            if not isinstance(cur_data[0], (list, tuple, dict)):
                cur_data = (cur_data,)

            for r, data in enumerate(cur_data, row):
                for _ in range(r - lines_count):  # 若行数不够，填充行数
                    lines.append([])
                    lines_count += 1
                row_num = r - 1
                lines[row_num], rewrite_header = self._header[None].make_change_list(lines[row_num],
                                                                                     data, col,
                                                                                     self._auto_new_header,
                                                                                     rewrite_header, 'csv')

                # if isinstance(data, dict):
                #     data = header.make_num_dict(data, 'csv')
                #     raw_data = {c: v for c, v in enumerate(lines[row_num], 1)}
                #     raw_data = {**raw_data, **data}
                #     lines[row_num] = [raw_data[c] for c in range(1, max(raw_data) + 1)]
                #
                # else:
                #     # 若列数不够，填充空列
                #     lines[row_num].extend([''] * (col - len(lines[row_num]) + len(data) - 1))
                #     for k, j in enumerate(data):  # 填充数据
                #         lines[row_num][col + k - 1] = process_content_str(j)

        if rewrite_header:
            for _ in range(self._header_row - lines_count):  # 若行数不够，填充行数
                lines.append([])
            lines[self._header_row - 1] = list(header.num_key.values())

        file.close()
        writer = csv_writer(open(self.path, 'w', encoding=self.encoding, newline=''),
                            delimiter=self.delimiter, quotechar=self.quote_char)
        writer.writerows(lines)

    def _to_txt(self):
        """记录数据到txt文件"""
        with open(self.path, 'a+', encoding=self.encoding) as f:
            all_data = [' '.join(ok_list_str(i)) for i in self._data]
            f.write('\n'.join(all_data) + '\n')

    def _to_jsonl(self):
        """记录数据到jsonl文件"""
        from json import dumps
        with open(self.path, 'a+', encoding=self.encoding) as f:
            all_data = [i if isinstance(i, str) else dumps(i) for i in self._data]
            f.write('\n'.join(all_data) + '\n')

    def _to_json(self):
        """记录数据到json文件"""
        from json import load, dump
        if self._file_exists or Path(self.path).exists():
            with open(self.path, 'r', encoding=self.encoding) as f:
                json_data = load(f)
        else:
            json_data = []

        for i in self._data:
            if isinstance(i, dict):
                for d in i:
                    i[d] = process_content_json(i[d])
                json_data.append(i)
            else:
                json_data.append([process_content_json(d) for d in i])

        self._file_exists = True
        with open(self.path, 'w', encoding=self.encoding) as f:
            dump(json_data, f)


def get_header(recorder, ws=None):
    """获取表头"""
    header = recorder._header.get(recorder._table, None)
    if header is not None:
        return header
    if not recorder.path or not Path(recorder.path).exists():
        return None

    if recorder.type == 'xlsx':
        if not ws:
            wb = load_workbook(recorder.path)
            if not recorder.table:
                ws = wb.active
            elif recorder.table not in wb.sheetnames:
                wb.close()
                return Header()
            else:
                ws = wb[recorder.table]
        recorder._header[recorder.table] = Header([i.value for i in ws[recorder._header_row]])

        if not ws:
            wb.close()
        return recorder._header[recorder.table]

    if recorder.type == 'csv':
        from csv import reader
        with open(recorder.path, 'r', newline='', encoding=recorder.encoding) as f:
            u = reader(f, delimiter=recorder.delimiter, quotechar=recorder.quote_char)
            try:
                for _ in range(recorder._header_row):
                    header = next(u)
            except StopIteration:  # 文件是空的
                header = []
        recorder._header[None] = Header(header)
        return recorder._header[None]


def new_sheet_fast(recorder, ws, data, first_wrote):
    """已有表头信息时向新表写入表头"""
    need_set_style = (recorder._header_row,)
    if recorder._header.get(ws.title, None) is not None and recorder._header_row > 0:
        for c, h in recorder._header[ws.title].items():
            ws.cell(row=recorder._header_row, column=c, value=h)

    elif (recorder._header_row > 0 and (data and (isinstance(data[0], dict) or (
            data[0] and isinstance(data[0], (list, tuple, set)) and isinstance(data[0][0], dict))))):
        # 第一个数据是dict，设置表头
        header = Header(list(data[0].keys()) if isinstance(data[0], dict) else list(data[0][0].keys()))
        recorder._header[ws.title] = header
        for c, h in header.items():
            ws.cell(row=recorder._header_row, column=c, value=h)

    elif data and recorder._header_row in (1, 0):  # 未设置表头，将第一个数据写到第一行
        first_wrote = True
        if data[0] and isinstance(data[0], (list, tuple)) and isinstance(data[0][0], (list, tuple)):
            for r, d in enumerate(data[0], 1):
                for c, v in enumerate(ok_list_xlsx(d), 1):
                    ws.cell(row=r, column=c, value=v)
            need_set_style = tuple(range(1, len(data[0]) + 1))
        else:  # 一维数据
            for c, v in enumerate(ok_list_xlsx(data[0]), 1):
                ws.cell(row=1, column=c, value=v)
        recorder._header[ws.title] = Header()

    else:
        recorder._header[ws.title] = Header()

    if recorder._styles or recorder._row_height:
        for r in need_set_style:
            set_style(recorder._row_height, recorder._styles, ws, r)

    return first_wrote


def new_sheet_slow(recorder, ws, data, style, first_data_wrote, first_style_wrote):
    """已有表头信息时向新表写入表头"""
    if recorder._header.get(ws.title, None) and recorder._header_row > 0:
        for c, h in recorder._header[ws.title].items():
            ws.cell(row=recorder._header_row, column=c, value=h)
        return first_data_wrote, first_style_wrote

    recorder._header[ws.title] = Header()
    if data:
        coord, data = data[0]
        if (recorder._header_row > 0 and ((isinstance(data, dict) or (
                data and isinstance(data, (list, tuple, set)) and isinstance(data[0], dict))))):
            # 第一个数据是dict，设置表头
            header = Header(data.keys() if isinstance(data, dict) else data[0].keys())
            recorder._header[ws.title] = header
            for c, h in header.items():
                ws.cell(row=recorder._header_row, column=c, value=h)

        elif (isinstance(coord, tuple) and coord[0]) or (isinstance(coord, str) and data[0][0]):
            # 有指定坐标
            return first_data_wrote, first_style_wrote

        elif recorder._header_row in (1, 0):  # 未设置表头，将第一个数据写到第一行
            first_data_wrote = True
            if isinstance(coord, tuple):
                if data and isinstance(data, (list, tuple)) and isinstance(data[0], (list, tuple)):
                    for r, d in enumerate(data, 1):
                        for c, v in enumerate(ok_list_xlsx(d), 1):
                            ws.cell(row=r, column=c, value=v)
                else:  # 一维数据
                    for c, v in enumerate(ok_list_xlsx(data), 1):
                        ws.cell(row=1, column=c, value=v)

            elif coord == 'set_link':
                set_link_to_ws(ws, data, True, recorder)

            elif coord == 'set_img':
                set_img_to_ws(ws, data, True, recorder)

        return first_data_wrote, first_style_wrote

    if style and isinstance(style[0][1][0], tuple) and not style[0][1][0][0]:
        first_style_wrote = True
        set_style_to_ws(ws, style[0], True, recorder, recorder._header.get(ws.title))
    return first_data_wrote, first_style_wrote


def get_xlsx_rows(recorder, header, key_cols, begin_row, sign_col, sign, deny_sign, count, ws):
    rows = ws.rows
    try:
        for _ in range(begin_row - 1):
            next(rows)
    except StopIteration:
        return []

    if sign_col is True or sign_col > ws.max_column:  # 获取所有行
        if count:
            rows = list(rows)[:count]  # todo: 是否要改进效率？
        if key_cols is True:  # 获取整行
            res = [header.make_row_data(ind, {col: cell.value for col, cell in enumerate(row, 1)})
                   for ind, row in enumerate(rows, begin_row)]
        else:  # 只获取对应的列
            res = [header.make_row_data(ind, {col: row[col - 1].value for col in key_cols})
                   for ind, row in enumerate(rows, begin_row)]

    else:  # 获取符合条件的行
        if count:
            res = handle_xlsx_rows_with_count(key_cols, deny_sign, header, rows,
                                              begin_row, sign_col, sign, count)
        else:
            res = handle_xlsx_rows_without_count(key_cols, deny_sign, header, rows, begin_row,
                                                 sign_col, sign)

    ws.parent.close()
    return res


def get_csv_rows(recorder, header, key_cols, begin_row, sign_col, sign, deny_sign, count, ws):
    sign = ['' if i is None else str(i) for i in sign]
    begin_row -= 1
    res = []
    with open(recorder.path, 'r', encoding=recorder.encoding) as f:
        reader = csv_reader(f, delimiter=recorder.delimiter, quotechar=recorder.quote_char)
        lines = list(reader)
        if not lines:
            return res

        if sign_col is True:  # 获取所有行
            header_len = len(header)
            for ind, line in enumerate(lines[begin_row:count + 1 if count else None], begin_row + 1):
                if key_cols is True:  # 获取整行
                    if not line:
                        res.append(header.make_row_data(ind, {col: '' for col in range(1, header_len + 1)}))
                    else:
                        line_len = len(line)
                        x = max(header_len, line_len)
                        res.append(header.make_row_data(ind, {col: line[col - 1] if col <= line_len else ''
                                                              for col in range(1, x + 1)}))
                else:  # 只获取对应的列
                    x = len(line) + 1
                    res.append(header.make_row_data(ind, {col: line[col - 1] if col < x else '' for col in key_cols}))

        else:  # 获取符合条件的行
            sign_col -= 1
            if count:
                handle_csv_rows_with_count(lines, begin_row, sign_col, sign, deny_sign,
                                           key_cols, res, header, count)
            else:
                handle_csv_rows_without_count(lines, begin_row, sign_col, sign, deny_sign,
                                              key_cols, res, header)

    return res


def handle_xlsx_rows_with_count(key_cols, deny_sign, header, rows, begin_row, sign_col, sign, count):
    got = 0
    res = []
    if key_cols is True:  # 获取整行
        if deny_sign:
            for ind, row in enumerate(rows, begin_row):
                if got == count:
                    break
                if row[sign_col - 1].value not in sign:
                    res.append(header.make_row_data(ind, {col: cell.value for col, cell in enumerate(row, 1)}))
                    got += 1
        else:
            for ind, row in enumerate(rows, begin_row):
                if got == count:
                    break
                if row[sign_col - 1].value in sign:
                    res.append(header.make_row_data(ind, {col: cell.value for col, cell in enumerate(row, 1)}))
                    got += 1

    else:  # 只获取对应的列
        if deny_sign:
            for ind, row in enumerate(rows, begin_row):
                if got == count:
                    break
                if row[sign_col - 1].value not in sign:
                    res.append(header.make_row_data(ind, {col: row[col - 1].value for col in key_cols}))
                    got += 1
        else:
            for ind, row in enumerate(rows, begin_row):
                if got == count:
                    break
                if row[sign_col - 1].value in sign:
                    res.append(header.make_row_data(ind, {col: row[col - 1].value for col in key_cols}))
                    got += 1
    return res


def handle_xlsx_rows_without_count(key_cols, deny_sign, header, rows, begin_row, sign_col, sign):
    if key_cols is True:  # 获取整行
        if deny_sign:
            return [header.make_row_data(ind, {col: cell.value for col, cell in enumerate(row, 1)})
                    for ind, row in enumerate(rows, begin_row)
                    if row[sign_col - 1].value not in sign]
        else:
            return [header.make_row_data(ind, {col: cell.value for col, cell in enumerate(row, 1)})
                    for ind, row in enumerate(rows, begin_row)
                    if row[sign_col - 1].value in sign]

    else:  # 只获取对应的列
        if deny_sign:
            return [header.make_row_data(ind, {col: row[col - 1].value for col in key_cols})
                    for ind, row in enumerate(rows, begin_row)
                    if row[sign_col - 1].value not in sign]
        else:
            return [header.make_row_data(ind, {col: row[col - 1].value for col in key_cols})
                    for ind, row in enumerate(rows, begin_row)
                    if row[sign_col - 1].value in sign]


def handle_csv_rows_with_count(lines, begin_row, sign_col, sign, deny_sign, key_cols, res, header, count):
    got = 0
    header_len = len(header)
    for ind, line in enumerate(lines[begin_row:], begin_row + 1):
        if got == count:
            break
        row_sign = '' if sign_col > len(line) - 1 else line[sign_col]
        if (row_sign not in sign) if deny_sign else (row_sign in sign):
            if key_cols is True:  # 获取整行
                if not line:
                    res.append(header.make_row_data(ind, {col: '' for col in range(1, header_len + 1)}))
                else:
                    line_len = len(line)
                    x = max(header_len, line_len)
                    res.append(header.make_row_data(ind, {col: line[col - 1] if col <= line_len else ''
                                                          for col in range(1, x + 1)}))
            else:  # 只获取对应的列
                x = len(line) + 1
                res.append(header.make_row_data(ind, {col: line[col - 1] if col < x else '' for col in key_cols}))
            got += 1


def handle_csv_rows_without_count(lines, begin_row, sign_col, sign, deny_sign, key_cols, res, header):
    header_len = len(header)
    for ind, line in enumerate(lines[begin_row:], begin_row + 1):
        row_sign = '' if sign_col > len(line) - 1 else line[sign_col]
        if (row_sign not in sign) if deny_sign else (row_sign in sign):
            if key_cols is True:  # 获取整行
                if not line:
                    res.append(header.make_row_data(ind, {col: '' for col in range(1, header_len + 1)}))
                else:
                    line_len = len(line)
                    x = max(header_len, line_len)
                    res.append(header.make_row_data(ind, {col: line[col - 1] if col <= line_len else ''
                                                          for col in range(1, x + 1)}))
            else:  # 只获取对应的列
                x = len(line) + 1
                res.append(header.make_row_data(ind, {col: line[col - 1] if col < x else '' for col in key_cols}))


def get_and_set_csv_header(recorder, new_csv, file, writer):
    if new_csv:
        if recorder._header[None] is None and recorder.data:
            if isinstance(recorder.data[0], dict):
                recorder._header[None] = Header([h for h in recorder.data[0].keys() if isinstance(h, str)])
            elif isinstance(recorder.data[0], (list, tuple)) and recorder.data[0] and isinstance(
                    recorder.data[0][0], dict):
                recorder._header[None] = Header([h for h in recorder.data[0][0].keys() if isinstance(h, str)])
            else:
                recorder._header[None] = Header()
        else:
            recorder._header[None] = Header()

        if recorder._header[None]:
            for _ in range(recorder._header_row - 1):
                writer.writerow([])
            writer.writerow(ok_list_str(recorder._header[None]))

    elif recorder._header[None] is None:
        file.seek(0)
        reader = csv_reader(file, delimiter=recorder.delimiter, quotechar=recorder.quote_char)
        header = []
        try:
            for _ in range(recorder._header_row):
                header = next(reader)
        except StopIteration:
            pass
        file.seek(2)
        recorder._header[None] = Header(header)


def data2ws_no_style(ws, header, row, col, data, not_new, max_row):
    for r, curr_data in enumerate(data, row):
        if isinstance(curr_data, dict):
            for c, val in header.make_num_dict(curr_data, 'xlsx').items():
                ws.cell(r, c, value=val)
        else:
            for key, j in enumerate(curr_data):
                ws.cell(r, col + key, value=process_content_xlsx(j))


def data2ws_has_style(ws, header, row, col, data, not_new, max_row):
    if not_new:  # 非新行
        styles = []
        for r, curr_data in enumerate(data, row):
            if isinstance(curr_data, dict):
                style = []
                for c, val in header.make_num_dict(curr_data, 'xlsx').items():
                    ws.cell(r, c, value=val)
                    style.append(c)
                styles.append(style)
            else:
                for key, j in enumerate(curr_data):
                    ws.cell(r, col + key, value=process_content_xlsx(j))
                styles.append(range(col, len(curr_data) + col))
        if row > 0 and max_row >= row - 1:
            copy_some_row_style(ws, row, styles)

    else:  # 新行，复制整行样式
        data2ws_no_style(ws, header, row, col, data, not_new, max_row)
        if row > 0 and max_row >= row - 1:
            copy_full_row_style(ws, row, data)


def set_link_to_ws(ws, data, empty, recorder):
    max_row = 0 if empty else ws.max_row
    coord = parse_coord(data[0], recorder.data_col)
    row, col = get_usable_coord(coord, max_row, ws)
    cell = ws.cell(row, col)
    has_link = True if cell.hyperlink else False
    cell.hyperlink = None if data[1] is None else process_content_str(data[1])
    if data[2] is not None:
        cell.value = process_content_xlsx(data[2])
    if data[1]:
        if recorder._link_style:
            recorder._link_style.to_cell(cell, replace=False)
    elif has_link:
        NoneStyle().to_cell(cell, replace=False)


def set_img_to_ws(ws, data, empty, recorder):
    max_row = 0 if empty else ws.max_row
    coord, img_path, width, height = data
    coord = parse_coord(coord, recorder.data_col)
    row, col = get_usable_coord(coord, max_row, ws)

    from openpyxl.drawing.image import Image
    img = Image(img_path)
    if width and height:
        img.width = width
        img.height = height
    elif width:
        img.height = int(img.height * (width / img.width))
        img.width = width
    elif height:
        img.width = int(img.width * (height / img.height))
        img.height = height
    # ws.add_image(img, (row, Header._NUM_KEY[col]))
    ws.add_image(img, f'{Header._NUM_KEY[col]}{row}')


def set_style_to_ws(ws, data, empty, recorder, header):
    """批量设置单元格格式到sheet"""
    if data[0] in ('replace_style', 'cover_style'):
        mode = data[0] == 'replace_style'
        coord = data[1][0]
        max_row = 0 if empty else ws.max_row

        if isinstance(data[1][1], dict):
            none_style = NoneStyle()
            coord = parse_coord(coord, recorder.data_col)
            row, col = get_usable_coord(coord, max_row, ws)
            for h, s in header.make_num_dict(data[1][1], None).items():
                if not s:
                    s = none_style
                s.to_cell(ws.cell(row, header[h]), replace=mode)
            return

        style = NoneStyle() if data[1][1] is None else data[1][1]
        if isinstance(coord, int) or (isinstance(coord, str) and coord.isdigit()):
            for c in ws[coord]:
                style.to_cell(c, replace=mode)

        elif isinstance(coord, str):
            if ':' in coord:
                for c in ws[coord]:
                    for cc in c:
                        style.to_cell(cc, replace=mode)
            elif coord.isdigit() or coord.isalpha():
                for c in ws[coord]:
                    style.to_cell(c, replace=mode)

        else:
            coord = parse_coord(coord, recorder.data_col)
            row, col = get_usable_coord(coord, max_row, ws)
            style.to_cell(ws.cell(row, col), replace=mode)

    elif data[0] == 'set_width':
        col, width = data[1]
        if isinstance(col, int):
            col = Header._NUM_KEY[col]
        for c in col.split(':'):
            if c.isdigit():
                c = Header._NUM_KEY[int(c)]
            ws.column_dimensions[c].width = width

    elif data[0] == 'set_height':
        max_row = 0 if empty else ws.max_row
        row, height = data[1]
        row = get_usable_coord((row, 1), max_row, ws)[0]
        if isinstance(row, int):
            ws.row_dimensions[row].height = height
        elif isinstance(row, str):
            for r in row.split(':'):
                ws.row_dimensions[int(r)].height = height


def set_style(height, styles, ws, row):
    if height is not None:
        ws.row_dimensions[row].height = height

    if styles:
        for c, s in enumerate(styles, start=1):
            if s:
                s.to_cell(ws.cell(row=row, column=c))


def copy_some_row_style(ws, row, styles):
    """复制上一行指定列样式到后续行中"""
    _row_styles = [CellStyleCopier(c) for c in ws[row - 1]]
    for r, i in enumerate(styles, row):
        for c in i:
            if _row_styles[c - 1]:
                _row_styles[c - 1].to_cell(ws.cell(row=r, column=c))


def copy_full_row_style(ws, row, cur_data):
    """复制上一行整行样式到新行中"""
    _row_styles = [CellStyleCopier(i) for i in ws[row - 1]]
    height = ws.row_dimensions[row - 1].height
    for r in range(row, len(cur_data) + 1):
        for k, s in enumerate(_row_styles, start=1):
            if s:
                s.to_cell(ws.cell(row=r, column=k))
        ws.row_dimensions[r].height = height


def copy_part_row_style(ws, row, cur_data, col):
    """复制上一行局部（连续）样式到后续行中"""
    _row_styles = [CellStyleCopier(i) for i in ws[row - 1]]
    for r, i in enumerate(cur_data, row):
        for c in range(len(i)):
            if _row_styles[c + col - 1]:
                _row_styles[c + col - 1].to_cell(ws.cell(row=r, column=c + col))

# def _set_line_style(height, style, ws, row, max_col):
#     if height is not None:
#         ws.row_dimensions[row].height = height
#     if style:
#         for i in range(1, max_col + 1):
#             style.to_cell(ws.cell(row=row, column=i))
