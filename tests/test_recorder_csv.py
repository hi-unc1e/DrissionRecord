# -*- coding:utf-8 -*-
"""Tests for Recorder class with CSV format."""
import csv
from pathlib import Path

import pytest

from DrissionRecord import Recorder


class TestRecorderCSV:
    """Test cases for Recorder with CSV format."""

    def test_create_recorder_with_path(self, temp_csv):
        """Test creating a Recorder with a file path."""
        r = Recorder(temp_csv)
        assert r.path == temp_csv
        assert r.type == 'csv'

    def test_create_recorder_without_path(self):
        """Test creating a Recorder without a file path."""
        r = Recorder()
        assert r.path is None

    def test_set_path_using_setter(self, temp_dir):
        """Test setting path using the fluent API."""
        r = Recorder()
        r.set.path(str(Path(temp_dir) / 'test.csv'))
        assert r.type == 'csv'

    def test_add_single_data(self, temp_csv):
        """Test adding a single data row."""
        r = Recorder(temp_csv)
        r.add_data((1, 2, 3))
        r.record()

        with open(temp_csv, 'r', encoding='utf-8', newline='') as f:
            reader = csv.reader(f)
            rows = list(reader)
            assert rows == [['1', '2', '3']]

    def test_add_multiple_data(self, temp_csv):
        """Test adding multiple data rows at once."""
        r = Recorder(temp_csv)
        data = ((1, 2, 3), (4, 5, 6), (7, 8, 9))
        r.add_data(data)
        r.record()

        with open(temp_csv, 'r', encoding='utf-8', newline='') as f:
            reader = csv.reader(f)
            rows = list(reader)
            assert rows == [['1', '2', '3'], ['4', '5', '6'], ['7', '8', '9']]

    def test_add_dict_data(self, temp_csv):
        """Test adding dictionary data."""
        r = Recorder(temp_csv)
        r.add_data({'name': 'Alice', 'age': 30, 'city': 'Beijing'})
        r.record()

        with open(temp_csv, 'r', encoding='utf-8', newline='') as f:
            reader = csv.reader(f)
            rows = list(reader)
            # Dict keys are used as header, values as data row
            assert len(rows) >= 1
            assert rows[-1] == ['Alice', '30', 'Beijing']

    def test_cache_size_auto_record(self, temp_csv):
        """Test automatic recording when cache size is reached."""
        r = Recorder(temp_csv, cache_size=3)

        # Add 3 items - should trigger auto record
        r.add_data((1, 2, 3))
        r.add_data((4, 5, 6))
        r.add_data((7, 8, 9))

        # File should exist now
        assert Path(temp_csv).exists()

    def test_delimiter(self, temp_dir):
        """Test custom delimiter."""
        file_path = str(Path(temp_dir) / 'test.csv')
        r = Recorder(file_path)
        r.set.delimiter(';')
        r.add_data((1, 2, 3))
        r.record()

        with open(file_path, 'r', encoding='utf-8', newline='') as f:
            content = f.read()
            assert '1;2;3' in content

    def test_quote_char(self, temp_dir):
        """Test custom quote character."""
        file_path = str(Path(temp_dir) / 'test.csv')
        r = Recorder(file_path)
        r.set.quote_char("'")
        r.add_data(('a', 'b,c', 'd'))
        r.record()

        with open(file_path, 'r', encoding='utf-8', newline='') as f:
            content = f.read()
            assert "'b,c'" in content

    def test_set_header(self, temp_csv):
        """Test setting a header."""
        r = Recorder(temp_csv)
        r.set.header(['ID', 'Name', 'Age'])
        r.add_data((1, 'Alice', 30))
        r.record()

        with open(temp_csv, 'r', encoding='utf-8', newline='') as f:
            reader = csv.reader(f)
            rows = list(reader)
            assert rows[0] == ['ID', 'Name', 'Age']

    def test_set_header_row(self, temp_csv):
        """Test setting header row number."""
        r = Recorder(temp_csv)
        r.set.header_row(2)
        r.set.header(['ID', 'Name', 'Age'])
        r.add_data((1, 'Alice', 30))
        r.record()

        with open(temp_csv, 'r', encoding='utf-8', newline='') as f:
            reader = csv.reader(f)
            rows = list(reader)
            assert len(rows) >= 2
            assert rows[1] == ['ID', 'Name', 'Age']

    def test_encoding(self, temp_dir):
        """Test different encoding."""
        file_path = str(Path(temp_dir) / 'test.csv')
        r = Recorder(file_path)
        r.set.encoding('gbk')
        r.add_data(('中文', '测试'))
        r.record()

        with open(file_path, 'r', encoding='gbk', newline='') as f:
            reader = csv.reader(f)
            rows = list(reader)
            assert rows == [['中文', '测试']]

    def test_clear_data(self, temp_csv):
        """Test clearing cached data."""
        r = Recorder(temp_csv, cache_size=100)
        r.add_data((1, 2, 3))
        r.clear()
        assert r._data_count == 0
        assert len(r.data) == 0

    def test_coordinate_append(self, temp_csv):
        """Test adding data with coordinates."""
        r = Recorder(temp_csv)
        r.set.data_col(1)
        r.add_data((1, 2, 3), coord=None)  # Append to end
        r.record()

        with open(temp_csv, 'r', encoding='utf-8', newline='') as f:
            reader = csv.reader(f)
            rows = list(reader)
            assert len(rows) == 1

    def test_before_and_after(self, temp_csv):
        """Test before and after data insertion."""
        r = Recorder(temp_csv)
        r.set.before(['prefix1', 'prefix2'])
        r.set.after(['suffix1', 'suffix2'])
        r.add_data(('main1', 'main2'))
        r.record()

        with open(temp_csv, 'r', encoding='utf-8', newline='') as f:
            reader = csv.reader(f)
            rows = list(reader)
            assert rows == [['prefix1', 'prefix2', 'main1', 'main2', 'suffix1', 'suffix2']]

    def test_append_to_existing_file(self, temp_csv):
        """Test appending data to an existing CSV file."""
        # Create initial file
        r = Recorder(temp_csv)
        r.add_data((1, 2, 3))
        r.record()

        # Append more data
        r2 = Recorder(temp_csv)
        r2.add_data((4, 5, 6))
        r2.record()

        with open(temp_csv, 'r', encoding='utf-8', newline='') as f:
            reader = csv.reader(f)
            rows = list(reader)
            assert rows == [['1', '2', '3'], ['4', '5', '6']]

    def test_empty_data(self, temp_csv):
        """Test handling empty data."""
        r = Recorder(temp_csv)
        r.add_data(())
        r.record()

        # File should be created even with empty data
        assert Path(temp_csv).exists()

    def test_mixed_data_types(self, temp_csv):
        """Test handling mixed data types."""
        r = Recorder(temp_csv)
        r.add_data((123, 'text', 3.14, None, True))
        r.record()

        with open(temp_csv, 'r', encoding='utf-8', newline='') as f:
            reader = csv.reader(f)
            rows = list(reader)
            assert rows == [['123', 'text', '3.14', '', 'True']]

    def test_show_msg_false(self, temp_csv, capsys):
        """Test disabling messages."""
        r = Recorder(temp_csv)
        r.set.show_msg(False)
        r.add_data((1, 2, 3))
        r.record()

        captured = capsys.readouterr()
        assert '开始写入文件' not in captured.out
