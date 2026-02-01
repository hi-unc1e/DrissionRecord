# -*- coding:utf-8 -*-
"""Test fixtures for DrissionRecord tests."""
import os
import tempfile
from pathlib import Path

import pytest


@pytest.fixture
def temp_dir():
    """Create a temporary directory for test files."""
    with tempfile.TemporaryDirectory() as tmpdir:
        yield tmpdir


@pytest.fixture
def temp_csv(temp_dir):
    """Create a temporary CSV file path."""
    return str(Path(temp_dir) / 'test.csv')


@pytest.fixture
def temp_xlsx(temp_dir):
    """Create a temporary XLSX file path."""
    return str(Path(temp_dir) / 'test.xlsx')


@pytest.fixture
def temp_json(temp_dir):
    """Create a temporary JSON file path."""
    return str(Path(temp_dir) / 'test.json')


@pytest.fixture
def temp_jsonl(temp_dir):
    """Create a temporary JSONL file path."""
    return str(Path(temp_dir) / 'test.jsonl')


@pytest.fixture
def temp_txt(temp_dir):
    """Create a temporary TXT file path."""
    return str(Path(temp_dir) / 'test.txt')


@pytest.fixture
def temp_db(temp_dir):
    """Create a temporary SQLite database file path."""
    return str(Path(temp_dir) / 'test.db')


@pytest.fixture
def temp_byte_file(temp_dir):
    """Create a temporary binary file path."""
    return str(Path(temp_dir) / 'test.bin')


@pytest.fixture
def backup_dir(temp_dir):
    """Create a backup directory path."""
    backup = Path(temp_dir) / 'backup'
    backup.mkdir()
    return str(backup)
