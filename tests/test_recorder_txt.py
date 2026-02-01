# -*- coding:utf-8 -*-
"""Tests for Recorder class with TXT format."""
from pathlib import Path

import pytest

from DrissionRecord import Recorder


class TestRecorderTXT:
    """Test cases for Recorder with TXT format."""

    def test_create_recorder_txt(self, temp_txt):
        """Test creating a Recorder for TXT format."""
        r = Recorder(temp_txt)
        assert r.type == 'txt'

    def test_add_data_txt(self, temp_txt):
        """Test adding data to TXT file."""
        r = Recorder(temp_txt)
        r.add_data((1, 2, 3))
        r.record()

        with open(temp_txt, 'r', encoding='utf-8') as f:
            content = f.read()
            assert '1 2 3' in content

    def test_add_multiple_rows_txt(self, temp_txt):
        """Test adding multiple rows to TXT file."""
        r = Recorder(temp_txt)
        data = ((1, 2, 3), (4, 5, 6), (7, 8, 9))
        r.add_data(data)
        r.record()

        with open(temp_txt, 'r', encoding='utf-8') as f:
            lines = f.readlines()
            assert len(lines) == 3
            assert '1 2 3\n' == lines[0]

    def test_string_data_txt(self, temp_txt):
        """Test adding string data to TXT file."""
        r = Recorder(temp_txt)
        r.add_data('hello world')
        r.record()

        with open(temp_txt, 'r', encoding='utf-8') as f:
            content = f.read()
            assert 'hello world' in content

    def test_dict_data_txt(self, temp_txt):
        """Test adding dict data to TXT file."""
        r = Recorder(temp_txt)
        r.add_data({'name': 'Alice', 'age': 30})
        r.record()

        with open(temp_txt, 'r', encoding='utf-8') as f:
            content = f.read()
            assert 'Alice' in content
            assert '30' in content

    def test_append_to_existing_txt(self, temp_txt):
        """Test appending to existing TXT file."""
        r = Recorder(temp_txt)
        r.add_data((1, 2, 3))
        r.record()

        r2 = Recorder(temp_txt)
        r2.add_data((4, 5, 6))
        r2.record()

        with open(temp_txt, 'r', encoding='utf-8') as f:
            lines = f.readlines()
            assert len(lines) == 2

    def test_encoding_txt(self, temp_dir):
        """Test different encoding for TXT files."""
        file_path = str(Path(temp_dir) / 'test.txt')
        r = Recorder(file_path)
        r.set.encoding('gbk')
        r.add_data(('中文测试',))
        r.record()

        with open(file_path, 'r', encoding='gbk') as f:
            content = f.read()
            assert '中文测试' in content

    def test_rows_method_txt(self, temp_txt):
        """Test the rows() method with TXT files."""
        r = Recorder(temp_txt)
        lines = ['line1', 'line2', 'line3', 'line4']
        for line in lines:
            r.add_data((line,))
        r.record()

        # Read all rows
        rows = r.rows()
        assert len(rows) == 4
        assert rows[0].value == 'line1'

        # Read with count
        rows = r.rows(count=2)
        assert len(rows) == 2

        # Read with begin_row
        rows = r.rows(begin_row=3)
        assert len(rows) == 2
        assert rows[0].value == 'line3'

    def test_empty_data_txt(self, temp_txt):
        """Test handling empty data in TXT file."""
        r = Recorder(temp_txt)
        r.add_data(())
        r.record()

        # File should be created
        assert Path(temp_txt).exists()

    def test_multiline_data_txt(self, temp_txt):
        """Test data with multiple elements in TXT format."""
        r = Recorder(temp_txt)
        r.add_data(('a', 'b', 'c', 'd'))
        r.record()

        with open(temp_txt, 'r', encoding='utf-8') as f:
            content = f.read().strip()
            assert content == 'a b c d'

    def test_numeric_data_txt(self, temp_txt):
        """Test numeric data in TXT format."""
        r = Recorder(temp_txt)
        r.add_data((123, 45.67, -89))
        r.record()

        with open(temp_txt, 'r', encoding='utf-8') as f:
            content = f.read()
            assert '123' in content
            assert '45.67' in content
            assert '-89' in content

    def test_none_values_txt(self, temp_txt):
        """Test handling None values in TXT format."""
        r = Recorder(temp_txt)
        r.add_data(('a', None, 'b'))
        r.record()

        with open(temp_txt, 'r', encoding='utf-8') as f:
            content = f.read()
            # None values are converted to empty strings
            assert 'a' in content
            assert 'b' in content
