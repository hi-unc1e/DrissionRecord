# -*- coding:utf-8 -*-
from pathlib import Path
from sqlite3 import connect
from time import sleep

from .base import BaseRecorder
from .setter import DBSetter
from .tools import ok_list_db, is_sigal_data, is_1D_data


class DBRecorder(BaseRecorder):
    def __init__(self, path=None, cache_size=1000, table=None):
        self._conn = None
        self._cur = None
        super().__init__(None, cache_size)
        if path:
            self.set.path(path, table)
        self._type = 'db'

    @property
    def set(self):
        if self._setter is None:
            self._setter = DBSetter(self)
        return self._setter

    @property
    def tables(self):
        self._connect()
        self._cur.execute("select name from sqlite_master where type='table'")
        tables = self._cur.fetchall()
        self._close_connection()
        return [i[0] for i in tables]

    def add_data(self, data, table=None):
        while self._pause_add:  # 等待其它线程写入结束
            sleep(.02)

        table = table or self.table
        if not isinstance(table, str):
            raise RuntimeError('未指定数据库表名。')

        data = self._handle_data(data)
        self._data.setdefault(table, []).extend(data)

        if 0 < self.cache_size <= self._data_count:
            self.record()

    def run_sql(self, sql, single=True, commit=False):
        self._connect()
        self._cur.execute(sql)
        r = self._cur.fetchone() if single else self._cur.fetchall()
        if commit:
            self._conn.commit()
        self._close_connection()
        return r

    def _connect(self):
        Path(self.path).parent.mkdir(parents=True, exist_ok=True)
        self._conn = connect(self.path)
        self._cur = self._conn.cursor()

    def _close_connection(self):
        if self._conn is not None:
            try:
                self._cur.close()
                self._conn.close()
            except:
                pass

    def _to_database(self, data_list, table, tables):
        if isinstance(data_list[0], dict):  # 检查是否要新增列
            keys = data_list[0].keys()
            for key in keys:
                if str(key) not in tables[table]:
                    if self._auto_new_header:
                        sql = f'ALTER TABLE `{table}` ADD COLUMN `{key}`'
                        self._cur.execute(sql)
                        tables[table].append(key)
                    else:
                        data_list[0].pop(key)

            question_masks = ','.join('?' * len(data_list[0]))
            keys_txt = '`' + '`,`'.join(data_list[0]) + '`'
            values = [ok_list_db(i.values()) for i in data_list]
            sql = f'INSERT INTO `{table}` ({keys_txt}) values ({question_masks})'

        else:
            question_masks = ','.join('?' * len(tables[table]))
            values = data_list
            sql = f'INSERT INTO `{table}` values ({question_masks})'

        self._cur.executemany(sql, values)

    def _record(self):
        self._connect()  # 获取所有表名和列名
        self._cur.execute("select name from sqlite_master where type='table'")
        tables = {}
        for table in self._cur.fetchall():
            self._cur.execute(f"PRAGMA table_info(`{table[0]}`)")
            tables[table[0]] = [i[1] for i in self._cur.fetchall()]

        for table, data in self._data.items():
            data_list = []
            if isinstance(data[0], dict):
                curr_keys = data[0].keys()
            else:
                curr_keys = len(data[0])

            for d in data:
                if isinstance(d, dict):
                    tmp_keys = d.keys()
                    if table not in tables:
                        keys = list(d.keys())
                        self._cur.execute(f"CREATE TABLE `{table}` (`{'`,`'.join(keys)}`)")
                        tables[table] = keys

                else:
                    if table not in tables:
                        self._close_connection()
                        raise TypeError('新建表格首次须接收数据需为dict格式。')
                    tmp_keys = len(d)
                    long = len(tables[table])
                    if long > tmp_keys:
                        d = ok_list_db(d)
                        d.extend([None] * (long - tmp_keys))
                    elif long < tmp_keys:
                        self._close_connection()
                        raise RuntimeError('数据个数大于列数（注意before和after属性）。')

                if tmp_keys != curr_keys:
                    self._to_database(data_list, table, tables)
                    curr_keys = tmp_keys
                    data_list = []

                data_list.append(d)

            if data_list:
                self._to_database(data_list, table, tables)

        self._conn.commit()
        self._close_connection()

    def _handle_data(self, data):
        if is_sigal_data(data):
            data = (self._make_final_data(self, (data,)),)
            self._data_count += 1
        elif not data:
            data = (self._make_final_data(self, tuple()),)
            self._data_count += 1
        elif is_1D_data(data):
            data = [self._make_final_data(self, data)]
            self._data_count += 1
        else:  # 二维数组
            data = [self._make_final_data(self, (d,)) if is_sigal_data(d)
                    else self._make_final_data(self, d) for d in data]
            self._data_count += len(data)
        return data
