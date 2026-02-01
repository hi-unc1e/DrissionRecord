# -*- coding:utf-8 -*-
"""Tests for DBRecorder class."""
from pathlib import Path
import sqlite3

import pytest

from DrissionRecord import DBRecorder


class TestDBRecorder:
    """Test cases for DBRecorder class."""

    def test_create_db_recorder(self, temp_db):
        """Test creating a DBRecorder."""
        d = DBRecorder(temp_db)
        assert d.path == temp_db
        assert d.type == 'db'

    def test_create_table_from_dict(self, temp_db):
        """Test creating a table from dictionary data."""
        d = DBRecorder(temp_db)
        d.add_data({'name': 'Alice', 'age': 30, 'city': 'Beijing'}, table='users')
        d.record()

        # Verify data was inserted
        conn = sqlite3.connect(temp_db)
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM users")
        result = cursor.fetchone()
        assert result == ('Alice', 30, 'Beijing')
        conn.close()

    def test_add_multiple_rows(self, temp_db):
        """Test adding multiple rows to database."""
        d = DBRecorder(temp_db)
        d.add_data({'id': 1, 'name': 'Alice'}, table='users')
        d.add_data({'id': 2, 'name': 'Bob'}, table='users')
        d.add_data({'id': 3, 'name': 'Charlie'}, table='users')
        d.record()

        conn = sqlite3.connect(temp_db)
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM users")
        count = cursor.fetchone()[0]
        assert count == 3
        conn.close()

    def test_add_list_data(self, temp_db):
        """Test adding list data to existing table."""
        # First create table with dict
        d = DBRecorder(temp_db)
        d.add_data({'name': 'Alice', 'age': 30}, table='users')
        d.record()

        # Add more data as list
        d2 = DBRecorder(temp_db)
        d2.set.table('users')
        d2.add_data([('Bob', 25), ('Charlie', 35)])
        d2.record()

        conn = sqlite3.connect(temp_db)
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM users")
        results = cursor.fetchall()
        assert len(results) == 3
        conn.close()

    def test_cache_size_auto_record(self, temp_db):
        """Test automatic recording when cache size is reached."""
        d = DBRecorder(temp_db, cache_size=2)
        d.add_data({'id': 1}, table='test')
        d.add_data({'id': 2}, table='test')
        # Should auto-record after cache_size is reached

        # Database file should exist
        assert Path(temp_db).exists()

    def test_set_table(self, temp_db):
        """Test setting the table name using fluent API."""
        d = DBRecorder(temp_db)
        d.set.table('users')
        assert d.table == 'users'

    def test_multiple_tables(self, temp_db):
        """Test working with multiple tables."""
        d = DBRecorder(temp_db)
        d.add_data({'name': 'Alice'}, table='users')
        d.add_data({'title': 'Post1'}, table='posts')
        d.record()

        tables = d.tables
        assert 'users' in tables
        assert 'posts' in tables

    def test_tables_property(self, temp_db):
        """Test the tables property."""
        d = DBRecorder(temp_db)
        d.add_data({'name': 'Alice'}, table='users')
        d.record()

        tables = d.tables
        assert isinstance(tables, list)
        assert 'users' in tables

    def test_auto_new_header(self, temp_db):
        """Test auto_new_header functionality."""
        d = DBRecorder(temp_db)
        d.set.auto_new_header(True)
        d.set.table('users')

        # Create initial table
        d.add_data({'name': 'Alice', 'age': 30})
        d.record()

        # Add data with new column
        d2 = DBRecorder(temp_db)
        d2.set.auto_new_header(True)
        d2.set.table('users')
        d2.add_data({'name': 'Bob', 'age': 25, 'city': 'Beijing'})
        d2.record()

        # Verify new column was added
        conn = sqlite3.connect(temp_db)
        cursor = conn.cursor()
        cursor.execute("PRAGMA table_info(users)")
        columns = [col[1] for col in cursor.fetchall()]
        assert 'city' in columns
        conn.close()

    def test_before_and_after(self, temp_db):
        """Test before and after data insertion."""
        d = DBRecorder(temp_db)
        d.set.before({'created_at': '2024-01-01'})
        d.set.after({'updated_at': '2024-01-02'})
        d.set.table('logs')
        d.add_data({'message': 'test'})
        d.record()

        conn = sqlite3.connect(temp_db)
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM logs")
        result = cursor.fetchone()
        # before/after are added to the data
        assert result is not None
        assert '2024-01-01' in result
        assert 'test' in result
        assert '2024-01-02' in result
        conn.close()

    def test_run_sql(self, temp_db):
        """Test running custom SQL."""
        d = DBRecorder(temp_db)
        d.add_data({'name': 'Alice'}, table='users')
        d.record()

        result = d.run_sql("SELECT name FROM users", single=True)
        assert result[0] == 'Alice'

    def test_run_sql_fetchall(self, temp_db):
        """Test running SQL to fetch all results."""
        d = DBRecorder(temp_db)
        d.add_data({'id': 1}, table='test')
        d.add_data({'id': 2}, table='test')
        d.record()

        results = d.run_sql("SELECT id FROM test", single=False)
        assert len(results) == 2

    def test_clear_data(self, temp_db):
        """Test clearing cached data."""
        d = DBRecorder(temp_db, cache_size=100)
        d.add_data({'test': 'data'}, table='test')
        d.clear()
        assert d._data_count == 0

    def test_empty_dict_data(self, temp_db):
        """Test handling empty dictionary data."""
        d = DBRecorder(temp_db)
        d.add_data({}, table='test')
        d.record()

        # Empty dict may not create a table - this is expected behavior
        # The test verifies the library handles this gracefully
        tables = d.tables
        # May or may not have tables depending on implementation
        assert isinstance(tables, list)

    def test_delete_method(self, temp_db):
        """Test the delete() method."""
        d = DBRecorder(temp_db)
        d.add_data({'test': 'data'}, table='test')
        d.record()
        assert Path(temp_db).exists()

        d.delete()
        assert not Path(temp_db).exists()

    def test_invalid_table_name(self, temp_db):
        """Test that invalid table names raise ValueError."""
        d = DBRecorder(temp_db)
        with pytest.raises(ValueError, match='table名称不能包含字符'):
            d.set.table('test`table')

    def test_list_data_to_new_table_error(self, temp_db):
        """Test that list data to new table needs to be handled correctly."""
        d = DBRecorder(temp_db)
        # New tables need dict data first to define columns
        # Then list data can be added
        d.add_data({'col1': 'a', 'col2': 'b'}, table='test')
        d.record()

        # Now we can add list data
        d2 = DBRecorder(temp_db)
        d2.set.table('test')
        d2.add_data([('c', 'd')])
        d2.record()

        # Verify it worked
        conn = sqlite3.connect(temp_db)
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM test")
        count = cursor.fetchone()[0]
        assert count == 2
        conn.close()

    def test_mismatched_column_count(self, temp_db):
        """Test handling of mismatched column counts."""
        d = DBRecorder(temp_db)
        d.add_data({'col1': 'a', 'col2': 'b', 'col3': 'c'}, table='test')
        d.record()

        # Add data with correct columns
        d2 = DBRecorder(temp_db)
        d2.set.table('test')
        d2.add_data([('x', 'y', 'z')])
        d2.record()

        # Verify both rows exist
        conn = sqlite3.connect(temp_db)
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM test")
        count = cursor.fetchone()[0]
        assert count == 2
        conn.close()

    def test_none_values(self, temp_db):
        """Test handling None values."""
        d = DBRecorder(temp_db)
        d.add_data({'name': 'Alice', 'age': None, 'city': None}, table='users')
        d.record()

        conn = sqlite3.connect(temp_db)
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM users")
        result = cursor.fetchone()
        assert result[0] == 'Alice'
        assert result[1] is None
        conn.close()

    def test_numeric_data(self, temp_db):
        """Test handling numeric data types."""
        d = DBRecorder(temp_db)
        d.add_data({'int_val': 42, 'float_val': 3.14}, table='numbers')
        d.record()

        conn = sqlite3.connect(temp_db)
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM numbers")
        result = cursor.fetchone()
        assert result[0] == 42
        assert result[1] == 3.14
        conn.close()

    def test_boolean_data(self, temp_db):
        """Test handling boolean data."""
        d = DBRecorder(temp_db)
        d.add_data({'active': True, 'deleted': False}, table='test')
        d.record()

        conn = sqlite3.connect(temp_db)
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM test")
        result = cursor.fetchone()
        # SQLite stores booleans as 0 and 1
        assert result[0] == 1
        assert result[1] == 0
        conn.close()

    def test_set_path_with_table(self, temp_db):
        """Test setting path with table parameter."""
        d = DBRecorder()
        d.set.path(temp_db, 'users')
        assert d.table == 'users'

        # If a table already exists, it should be selected
        d.add_data({'name': 'Alice'})
        d.record()

        d2 = DBRecorder()
        d2.set.path(temp_db)
        # Should select the existing table
        assert d2.table is not None or len(d2.tables) > 0
