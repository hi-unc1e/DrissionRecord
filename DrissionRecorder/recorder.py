# -*- coding:utf-8 -*-
from csv import reader as csv_reader, writer as csv_writer
from pathlib import Path
from time import sleep

from openpyxl.reader.excel import load_workbook

from .base import BaseRecorder
from .cell_style import CellStyleCopier, CellStyle, NoneStyle
from .setter import RecorderSetter, set_csv_header
from .tools import (ok_list_str, process_content_xlsx, process_content_json, process_content_str,
                    get_key_cols, Header, get_wb, get_ws, get_csv, parse_coord, do_nothing,
                    get_usable_coord, get_real_coord, is_sigal_data, is_1D_data, get_real_col)


class Recorder(BaseRecorder):
    _SUPPORTS = ('csv', 'xlsx', 'txt', 'jsonl', 'json')

    def __init__(self, path=None, cache_size=1000):
        self._header = {None: None}
        self._methods = {'xlsx': self._to_xlsx_fast,
                         'csv': self._to_csv_fast,
                         'txt': self._to_txt,
                         'jsonl': self._to_jsonl,
                         'json': self._to_json,
                         'set_img': do_nothing,
                         'set_link': do_nothing,
                         'set_style': do_nothing,
                         'set_row_height': do_nothing,
                         'set_col_width': do_nothing,
                         }
        super().__init__(path=path, cache_size=cache_size)
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
        if self._setter is None:
            self._setter = RecorderSetter(self)
        return self._setter

    @property
    def delimiter(self):
        return self._delimiter

    @property
    def quote_char(self):
        return self._quote_char

    @property
    def header(self):
        if self.type not in ('csv', 'xlsx'):
            raise TypeError('header属性只支持csv和xlsx类型文件。')
        return get_header(self)

    def add_data(self, data, coord=None, table=None):
        coord = parse_coord(coord, self.data_col)
        data, data_len = self._handle_data(data, coord)
        self._add(data, table,
                  True if self._fast and coord[0] and self._type in ('csv', 'xlsx') else False,
                  data_len)

    def _set_link(self, coord, link, content=None, table=None):
        self._add([{'type': 'set_link', 'link': link, 'content': content,
                    'coord': parse_coord(coord, self.data_col)}], table, self._fast, 1)

    def _set_img(self, coord, img_path, width=None, height=None, table=None):
        self._add([{'type': 'set_link', 'img_path': img_path, 'width': width, 'height': height,
                    'coord': parse_coord(coord, self.data_col)}], table, self._fast, 1)

    def _set_style(self, coord, style, replace=True, table=None):
        self._add([{'type': 'replace_style' if replace else 'cover_style',
                    'style': style, 'coord': parse_coord(coord, self.data_col)}], table, self._fast, 1)

    def _set_row_height(self, row, height, table=None):
        if not row:
            raise ValueError('row不能为0或None。')
        self._add([{'type': 'set_height', 'row': row, 'height': height}], table, self._fast, 1)

    def _set_col_width(self, col, width, table=None):
        if not col:
            raise ValueError('col不能为0或None。')
        self._add([{'type': 'set_width', 'col': col, 'width': width}], table, self._fast, 1)

    def _add(self, data, table, to_slow, num):
        while self._pause_add:  # 等待其它线程写入结束
            sleep(.02)

        if to_slow:
            self._slow_mode()

        if self._type == 'xlsx':
            if table is None:
                table = self._table
            elif table is True:
                table = None
            self._data.setdefault(table, []).extend(data)
        elif self._type:
            self._data.extend(data)
        else:
            raise RuntimeError('请设置文件路径。')

        self._data_count += num
        if 0 < self.cache_size <= self._data_count:
            self.record()

    def set_link(self, coord, link, content=None, table=None):
        self._methods['set_link'](coord, link, content, table)

    def set_img(self, coord, img_path, width=None, height=None, table=None):
        self._methods['set_img'](coord, img_path, width, height, table)

    def set_style(self, coord, style, replace=True, table=None):
        self._methods['set_style'](coord, style, replace, table)

    def set_row_height(self, row, height, table=None):
        self._methods['set_row_height'](row, height, table)

    def set_col_width(self, col, width, table=None):
        self._methods['set_col_width'](col, width, table)

    def rows(self, key_cols=True, sign_col=True, is_header=False,
             signs=None, deny_sign=False, count=None, begin_row=None):
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

    def _handle_data(self, data, coord):
        if is_sigal_data(data):
            data = [{'type': 'data', 'data': self._handle_data_method(self, (data,)), 'coord': coord}]
            data_len = 1
        elif not data:
            data = [{'type': 'data', 'data': self._handle_data_method(self, tuple()), 'coord': coord}]
            data_len = 1
        elif is_1D_data(data):
            data = [{'type': 'data', 'data': self._handle_data_method(self, data), 'coord': coord}]
            data_len = 1
        else:  # 二维数组
            if not coord[0] and self._fast:
                data = [{'type': 'data', 'coord': coord,
                         'data': (self._handle_data_method(self, (d,)) if is_sigal_data(d)
                                  else self._handle_data_method(self, d))}
                        for d in data]
                data_len = len(data)
            else:
                data = [self._handle_data_method(self, (d,)) if is_sigal_data(d)
                        else self._handle_data_method(self, d) for d in data]
                data = [{'type': '2Ddata', 'data': data, 'coord': coord}]
                data_len = 1
        return data, data_len

    def _record(self):
        self._methods[self.type]()
        if not self._fast:
            self._fast_mode()

    def _fast_mode(self):
        self._methods['xlsx'] = self._to_xlsx_fast
        self._methods['csv'] = self._to_csv_fast
        self._fast = True

    def _slow_mode(self):
        self._methods['xlsx'] = self._to_xlsx_slow
        self._methods['csv'] = self._to_csv_slow
        self._fast = False

    def _to_xlsx_fast(self):
        wb, new_file = get_wb(self)
        tables = wb.sheetnames
        rewrite_method = 'make_num_dict_rewrite' if self._auto_new_header else 'make_num_dict'

        for table, data in self._data.items():
            _row_styles = None
            _row_height = None
            ws, new_sheet = get_ws(wb, table, tables, new_file)
            new_file = False
            if table is None and ws.title not in self._header:
                self._header[ws.title] = self._header[None]

            if new_sheet:
                begin_row = handle_new_sheet(self, ws, data)
            elif self._header.get(ws.title, None) is None:
                self._header[ws.title] = Header([c.value for c in ws[self._header_row]])
                begin_row = ws.max_row
            else:
                begin_row = ws.max_row

            header = self._header[ws.title]
            if self._follow_styles and begin_row > 0:
                _row_styles = [CellStyleCopier(i) for i in ws[begin_row]]
                _row_height = ws.row_dimensions[begin_row].height

            begin_row += 1
            rewrite = False
            for row, d in enumerate(data, begin_row):
                rewrite = line2ws(ws, header, row,
                                  (get_real_col(d['coord'][1], len(header) if self._header_row > 0 else 1)
                                   if d['coord'][1] < 1 else d['coord'][1]),
                                  d['data'], rewrite_method, rewrite)

            if rewrite:
                for c in range(1, ws.max_column + 1):
                    ws.cell(self._header_row, c, value=header[c])

            if self._follow_styles and begin_row > 0:
                for r in range(begin_row, ws.max_row + 1):
                    set_style(_row_height, _row_styles, ws, r)

            elif self._styles or self._row_height:
                if isinstance(self._styles, dict):
                    styles = header.make_num_dict(self._styles, None)[0]
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
        wb, new_file = get_wb(self)
        tables = wb.sheetnames
        method = data2ws_has_style if self._follow_styles else data2ws_no_style
        rewrite_method = 'make_num_dict_rewrite' if self._auto_new_header else 'make_num_dict'

        for table, data in self._data.items():
            ws, new_sheet = get_ws(wb, table, tables, new_file)
            new_file = False
            if table is None and ws.title not in self._header:
                self._header[ws.title] = self._header[None]

            if new_sheet:
                begin_row = handle_new_sheet(self, ws, data)
            elif self._header.get(ws.title, None) is None:
                self._header[ws.title] = Header([c.value for c in ws[self._header_row]])
                begin_row = ws.max_row
            else:
                begin_row = ws.max_row

            header = self._header[ws.title]
            rewrite = False
            # --------------- 这里继续 ----------------
            if not begin_row and not self._data[table][0][0][0]:
                rewrite = data2ws_no_style(ws, header, begin_row, 1, self._data[table][0], False, 1,
                                           rewrite_method, False)
                data = self._data[table][1:]
            else:
                data = self._data[table]

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
                    rewrite = method(ws, header, row, col, cur_data, not_new, max_row, rewrite_method, rewrite)

            if rewrite:
                for c in range(1, ws.max_column + 1):
                    ws.cell(self._header_row, c, value=header[c])

        wb.save(self.path)
        wb.close()

    def _to_csv_fast(self):
        file, new_csv = get_csv(self)
        writer = csv_writer(file, delimiter=self.delimiter, quotechar=self.quote_char)
        get_and_set_csv_header(self, new_csv, file, writer)
        rewrite_method = 'make_insert_list_rewrite' if self._auto_new_header else 'make_insert_list'

        rewrite = False
        header = self._header[None]
        for d in self._data:
            i, rewrite = header.__getattribute__(rewrite_method)(d['data'], 'csv', rewrite)
            if d['coord'][1] != 1:
                col = get_real_col(d['coord'][1], len(header) if self._header_row > 0 else 1)
                i = [None] * (col - 1) + i
            writer.writerow(i)
        file.close()

        if rewrite:
            set_csv_header(self, self._header[None], self._header_row)

    def _to_csv_slow(self):
        file, new_csv = get_csv(self)
        writer = csv_writer(file, delimiter=self.delimiter, quotechar=self.quote_char)
        get_and_set_csv_header(self, new_csv, file, writer)
        file.seek(0)
        reader = csv_reader(file, delimiter=self.delimiter, quotechar=self.quote_char)
        lines = list(reader)
        lines_count = len(lines)
        header = self._header[None]

        rewrite = False
        method = 'make_change_list_rewrite' if self._auto_new_header else 'make_change_list'
        for i in self._data:
            data = i['data'] if i['type'] == '2Ddata' else (i['data'],)
            row, col = get_real_coord(i['coord'], lines_count, len(header) if self._header_row > 0 else 1)
            for r, da in enumerate(data, row):
                add_rows = r - lines_count
                if add_rows > 0:  # 若行数不够，填充行数
                    [lines.append([]) for _ in range(add_rows)]
                    lines_count += add_rows
                row_num = r - 1
                lines[row_num], rewrite = self._header[None].__getattribute__(method)(lines[row_num], da, col,
                                                                                      'csv', rewrite)

        if rewrite:
            [lines.append([]) for _ in range(self._header_row - lines_count)]  # 若行数不够，填充行数
            lines[self._header_row - 1] = list(header.num_key.values())

        file.close()
        writer = csv_writer(open(self.path, 'w', encoding=self.encoding, newline=''),
                            delimiter=self.delimiter, quotechar=self.quote_char)
        writer.writerows(lines)

    def _to_txt(self):
        with open(self.path, 'a+', encoding=self.encoding) as f:
            all_data = [' '.join(ok_list_str(i['data'])) for i in self._data]
            f.write('\n'.join(all_data) + '\n')

    def _to_jsonl(self):
        from json import dumps
        with open(self.path, 'a+', encoding=self.encoding) as f:
            all_data = [i['data'] if isinstance(i['data'], str) else dumps(i['data']) for i in self._data]
            f.write('\n'.join(all_data) + '\n')

    def _to_json(self):
        from json import load, dump
        if self._file_exists or Path(self.path).exists():
            with open(self.path, 'r', encoding=self.encoding) as f:
                json_data = load(f)
        else:
            json_data = []

        for i in self._data:
            i = i['data']
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


def handle_new_sheet(recorder, ws, data):
    if not recorder._header_row:
        return 0

    if recorder._header.get(ws.title, None) is not None:
        for c, h in recorder._header[ws.title].items():
            ws.cell(row=recorder._header_row, column=c, value=h)
        begin_row = recorder._header_row

    else:
        data = get_first_dict(data)
        if data:
            header = Header([h for h in recorder.data.keys() if isinstance(h, str)])
            recorder._header[ws.title] = header
            for c, h in header.items():
                ws.cell(row=recorder._header_row, column=c, value=h)
            begin_row = recorder._header_row

        else:
            recorder._header[ws.title] = Header()
            begin_row = 0

    if begin_row and (recorder._styles or recorder._row_height):
        set_style(recorder._row_height, recorder._styles, ws, recorder._header_row)

    return begin_row


def get_first_dict(data):
    if not data:
        return False
    elif data[0]['type'] == 'data' and isinstance(data[0]['data'], dict):
        return data[0]['data']
    elif data[0]['type'] == '2Ddata' and data[0]['data'] and isinstance(data[0]['data'][0], dict):
        return data[0]['data'][0]['data']


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
    if not recorder._header_row:
        return

    if new_csv:
        if recorder._header[None]:
            for _ in range(recorder._header_row - 1):
                writer.writerow([])
            writer.writerow(ok_list_str(recorder._header[None]))

        if recorder._header[None] is None and recorder.data:
            data = get_first_dict(recorder._data)
            if data:
                recorder._header[None] = Header([h for h in recorder.data.keys() if isinstance(h, str)])
            else:
                recorder._header[None] = Header()
        else:
            recorder._header[None] = Header()

    elif recorder._header[None] is None:  # 从文件读取表头
        file.seek(0)
        reader = csv_reader(file, delimiter=recorder.delimiter, quotechar=recorder.quote_char)
        header = []
        try:
            for _ in range(recorder._header_row):
                header = next(reader)
        except StopIteration:
            pass
        recorder._header[None] = Header(header)
        file.seek(2)


def line2ws(ws, header, row, col, data, rewrite_method, rewrite):
    if isinstance(data, dict):
        data, rewrite, header_len = header.__getattribute__(rewrite_method)(data, 'xlsx', rewrite)
        for c, val in data.items():
            ws.cell(row, c, value=process_content_xlsx(val))
    else:
        if col < 0:
            col = ws.max_column + col + 1
            if col < 0:
                col = 1
        for key, val in enumerate(data):
            ws.cell(row, col + key, value=process_content_xlsx(val))
    return rewrite


def data2ws_no_style(ws, header, row, col, data, not_new, max_row, rewrite_method, rewrite):
    for r, curr_data in enumerate(data, row):
        curr_data = curr_data['data']
        if isinstance(curr_data, dict):
            curr_data, rewrite, header_len = header.__getattribute__(rewrite_method)(curr_data, 'xlsx', rewrite)
            for c, val in curr_data.items():
                ws.cell(r, c, value=process_content_xlsx(val))
        else:
            for key, val in enumerate(curr_data):
                ws.cell(r, col + key, value=process_content_xlsx(val))
    return rewrite


def data2ws_has_style(ws, header, row, col, data, not_new, max_row, rewrite_method, rewrite):
    if not_new:  # 非新行
        styles = []
        for r, curr_data in enumerate(data, row):
            if isinstance(curr_data, dict):
                style = []
                curr_data, rewrite, header_len = header.__getattribute__(rewrite_method)(curr_data, 'xlsx', rewrite)
                for c, val in curr_data.items():
                    ws.cell(r, c, value=process_content_xlsx(val))
                    style.append(c)
                styles.append(style)
            else:
                for key, val in enumerate(curr_data):
                    ws.cell(r, col + key, value=process_content_xlsx(val))
                styles.append(range(col, len(curr_data) + col))
        if row > 0 and max_row >= row - 1:
            copy_some_row_style(ws, row, styles)

    else:  # 新行，复制整行样式
        data2ws_no_style(ws, header, row, col, data, not_new, max_row, rewrite_method, rewrite)
        if row > 0 and max_row >= row - 1:
            copy_full_row_style(ws, row, data)

    return rewrite


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
    if data[0] in ('replace_style', 'cover_style'):
        mode = data[0] == 'replace_style'
        coord = data[1][0]
        max_row = 0 if empty else ws.max_row

        if isinstance(data[1][1], dict):
            none_style = NoneStyle()
            coord = parse_coord(coord, recorder.data_col)
            row, col = get_usable_coord(coord, max_row, ws)
            for h, s in header.make_num_dict(data[1][1], None)[0].items():
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
    _row_styles = [CellStyleCopier(c) for c in ws[row - 1]]
    for r, i in enumerate(styles, row):
        for c in i:
            if _row_styles[c - 1]:
                _row_styles[c - 1].to_cell(ws.cell(row=r, column=c))


def copy_full_row_style(ws, row, cur_data):
    _row_styles = [CellStyleCopier(i) for i in ws[row - 1]]
    height = ws.row_dimensions[row - 1].height
    for r in range(row, len(cur_data) + 1):
        for k, s in enumerate(_row_styles, start=1):
            if s:
                s.to_cell(ws.cell(row=r, column=k))
        ws.row_dimensions[r].height = height


def copy_part_row_style(ws, row, cur_data, col):
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
