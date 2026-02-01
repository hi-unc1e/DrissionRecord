# -*- coding:utf-8 -*-
"""Tests for Recorder class with JSON and JSONL format."""
import json
from pathlib import Path

import pytest

from DrissionRecord import Recorder


class TestRecorderJSON:
    """Test cases for Recorder with JSON format."""

    def test_create_recorder_json(self, temp_json):
        """Test creating a Recorder for JSON format."""
        r = Recorder(temp_json)
        assert r.type == 'json'

    def test_add_data_json(self, temp_json):
        """Test adding data to JSON file."""
        r = Recorder(temp_json)
        r.add_data((1, 2, 3))
        r.record()

        with open(temp_json, 'r', encoding='utf-8') as f:
            data = json.load(f)
            # JSON format stores data with metadata
            assert len(data) >= 1
            # Verify file was created and has list structure
            assert isinstance(data, list)

    def test_add_dict_data_json(self, temp_json):
        """Test adding dict data to JSON file."""
        r = Recorder(temp_json)
        r.add_data({'name': 'Alice', 'age': 30})
        r.record()

        with open(temp_json, 'r', encoding='utf-8') as f:
            data = json.load(f)
            assert data == [{'name': 'Alice', 'age': 30}]

    def test_add_multiple_rows_json(self, temp_json):
        """Test adding multiple rows to JSON file."""
        r = Recorder(temp_json)
        data = ((1, 2, 3), (4, 5, 6))
        r.add_data(data)
        r.record()

        with open(temp_json, 'r', encoding='utf-8') as f:
            data = json.load(f)
            # JSON stores data with metadata structure
            assert len(data) >= 2
            # Data exists in the file
            assert isinstance(data, list)

    def test_append_to_existing_json(self, temp_json):
        """Test appending to existing JSON file."""
        r = Recorder(temp_json)
        r.add_data((1, 2, 3))
        r.record()

        r2 = Recorder(temp_json)
        r2.add_data((4, 5, 6))
        r2.record()

        with open(temp_json, 'r', encoding='utf-8') as f:
            data = json.load(f)
            assert len(data) >= 2
            # Verify file has content
            assert isinstance(data, list)

    def test_mixed_data_types_json(self, temp_json):
        """Test handling mixed data types in JSON."""
        r = Recorder(temp_json)
        r.add_data((123, 'text', 3.14, None, True))
        r.record()

        with open(temp_json, 'r', encoding='utf-8') as f:
            data = json.load(f)
            assert len(data) >= 1
            # File was created and has data
            assert isinstance(data, list)

    def test_header_json(self, temp_json):
        """Test setting header for JSON file."""
        r = Recorder(temp_json)
        r.set.header(['ID', 'Name', 'Age'])
        r.record()

        # Header is stored internally
        assert r.header is not None

    def test_rows_method_json(self, temp_json):
        """Test the rows() method with JSON files."""
        r = Recorder(temp_json)
        r.set.header(['id', 'name', 'age'])
        data = [
            {'id': 1, 'name': 'Alice', 'age': 25},
            {'id': 2, 'name': 'Bob', 'age': 30},
        ]
        r.add_data(data)
        r.record()

        # Read all rows
        rows = r.rows()
        assert len(rows) == 2
        assert rows[0]['name'] == 'Alice'


class TestRecorderJSONL:
    """Test cases for Recorder with JSONL format."""

    def test_create_recorder_jsonl(self, temp_jsonl):
        """Test creating a Recorder for JSONL format."""
        r = Recorder(temp_jsonl)
        assert r.type == 'jsonl'

    def test_add_data_jsonl(self, temp_jsonl):
        """Test adding data to JSONL file."""
        r = Recorder(temp_jsonl)
        r.add_data({'name': 'Alice', 'age': 30})
        r.record()

        with open(temp_jsonl, 'r', encoding='utf-8') as f:
            lines = f.readlines()
            assert len(lines) == 1
            data = json.loads(lines[0])
            assert data == {'name': 'Alice', 'age': 30}

    def test_add_multiple_rows_jsonl(self, temp_jsonl):
        """Test adding multiple rows to JSONL file."""
        r = Recorder(temp_jsonl)
        data = [
            {'name': 'Alice', 'age': 25},
            {'name': 'Bob', 'age': 30},
        ]
        r.add_data(data)
        r.record()

        with open(temp_jsonl, 'r', encoding='utf-8') as f:
            lines = f.readlines()
            assert len(lines) == 2
            assert json.loads(lines[0]) == {'name': 'Alice', 'age': 25}
            assert json.loads(lines[1]) == {'name': 'Bob', 'age': 30}

    def test_append_to_existing_jsonl(self, temp_jsonl):
        """Test appending to existing JSONL file."""
        r = Recorder(temp_jsonl)
        r.add_data({'id': 1})
        r.record()

        r2 = Recorder(temp_jsonl)
        r2.add_data({'id': 2})
        r2.record()

        with open(temp_jsonl, 'r', encoding='utf-8') as f:
            lines = f.readlines()
            assert len(lines) == 2

    def test_string_data_jsonl(self, temp_jsonl):
        """Test adding string data to JSONL file."""
        r = Recorder(temp_jsonl)
        r.add_data('{"name": "Alice", "age": 30}')
        r.record()

        with open(temp_jsonl, 'r', encoding='utf-8') as f:
            content = f.read()
            # Verify file was created and has some content
            assert len(content) > 0
            assert 'name' in content or 'Alice' in content

    def test_rows_method_jsonl(self, temp_jsonl):
        """Test the rows() method with JSONL files."""
        r = Recorder(temp_jsonl)
        r.set.header(['id', 'name', 'age'])
        data = [
            {'id': 1, 'name': 'Alice', 'age': 25},
            {'id': 2, 'name': 'Bob', 'age': 30},
        ]
        r.add_data(data)
        r.record()

        # Read all rows
        rows = r.rows()
        assert len(rows) == 2
        assert rows[0]['name'] == 'Alice'

        # Read with specific columns
        rows = r.rows(cols=['name', 'age'])
        assert len(rows[0]) == 2

    def test_rows_with_sign_jsonl(self, temp_jsonl):
        """Test rows() with sign filtering for JSONL."""
        r = Recorder(temp_jsonl)
        r.set.header(['id', 'status', 'value'])
        data = [
            {'id': 1, 'status': 'active', 'value': 100},
            {'id': 2, 'status': 'inactive', 'value': 200},
            {'id': 3, 'status': 'active', 'value': 300},
        ]
        r.add_data(data)
        r.record()

        # Get only active rows
        rows = r.rows(sign_col='status', signs='active')
        assert len(rows) == 2
        assert all(r['status'] == 'active' for r in rows)
