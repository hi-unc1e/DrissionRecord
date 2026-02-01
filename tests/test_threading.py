# -*- coding:utf-8 -*-
"""Tests for threading safety of recorders."""
import csv
import threading
from pathlib import Path
from time import sleep

import pytest

from DrissionRecord import Recorder, ByteRecorder, DBRecorder


@pytest.mark.threading
class TestThreadingRecorder:
    """Test cases for Recorder thread safety."""

    def test_concurrent_add_data(self, temp_csv):
        """Test multiple threads adding data concurrently."""
        r = Recorder(temp_csv, cache_size=50)

        def add_data(thread_id):
            for i in range(10):
                r.add_data((thread_id, i, f'data_{thread_id}_{i}'))

        threads = []
        for i in range(5):
            t = threading.Thread(target=add_data, args=(i,))
            threads.append(t)
            t.start()

        for t in threads:
            t.join()

        r.record()

        # Verify all data was written
        with open(temp_csv, 'r', encoding='utf-8', newline='') as f:
            reader = csv.reader(f)
            rows = list(reader)
            assert len(rows) == 50  # 5 threads * 10 rows

    def test_concurrent_add_and_record(self, temp_csv):
        """Test concurrent add_data and record calls."""
        r = Recorder(temp_csv, cache_size=1000)

        def adder():
            for i in range(20):
                r.add_data((i, 'data'))

        def recorder():
            for _ in range(5):
                sleep(0.01)
                r.record()

        threads = []
        for _ in range(2):
            threads.append(threading.Thread(target=adder))
        threads.append(threading.Thread(target=recorder))

        for t in threads:
            t.start()

        for t in threads:
            t.join()

        # File should exist and have data
        assert Path(temp_csv).exists()

    def test_concurrent_different_recorders_same_file(self, temp_csv):
        """Test multiple Recorder objects writing to same file."""
        def writer(thread_id):
            r = Recorder(temp_csv)
            for i in range(5):
                r.add_data((thread_id, i))
            r.record()

        threads = []
        for i in range(3):
            t = threading.Thread(target=writer, args=(i,))
            threads.append(t)
            t.start()

        for t in threads:
            t.join()

        # Verify data was written
        with open(temp_csv, 'r', encoding='utf-8', newline='') as f:
            reader = csv.reader(f)
            rows = list(reader)
            # Should have at least some rows
            assert len(rows) > 0

    def test_concurrent_xlsx_writing(self, temp_xlsx):
        """Test multiple threads writing to XLSX file."""
        r = Recorder(temp_xlsx, cache_size=100)

        def add_data(thread_id):
            for i in range(10):
                r.add_data((thread_id, i, f'thread_{thread_id}'))

        threads = []
        for i in range(3):
            t = threading.Thread(target=add_data, args=(i,))
            threads.append(t)
            t.start()

        for t in threads:
            t.join()

        r.record()

        # Verify file was created
        assert Path(temp_xlsx).exists()

    def test_concurrent_json_writing(self, temp_json):
        """Test multiple threads writing to JSON file."""
        r = Recorder(temp_json, cache_size=50)

        def add_data(thread_id):
            for i in range(5):
                r.add_data({'thread': thread_id, 'index': i})

        threads = []
        for i in range(4):
            t = threading.Thread(target=add_data, args=(i,))
            threads.append(t)
            t.start()

        for t in threads:
            t.join()

        r.record()

        # Verify file was created
        assert Path(temp_json).exists()


@pytest.mark.threading
class TestThreadingByteRecorder:
    """Test cases for ByteRecorder thread safety."""

    def test_concurrent_binary_write(self, temp_byte_file):
        """Test multiple threads writing binary data."""
        b = ByteRecorder(temp_byte_file, cache_size=50)

        def add_data(thread_id):
            for i in range(10):
                b.add_data(f'{thread_id}-{i}-'.encode())

        threads = []
        for i in range(5):
            t = threading.Thread(target=add_data, args=(i,))
            threads.append(t)
            t.start()

        for t in threads:
            t.join()

        b.record()

        # Verify file was created
        assert Path(temp_byte_file).exists()

    def test_concurrent_with_seek(self, temp_byte_file):
        """Test concurrent writes with seek operations."""
        b = ByteRecorder(temp_byte_file, cache_size=100)

        def writer(thread_id):
            for i in range(5):
                position = thread_id * 100 + i * 10
                b.add_data(b'TEST', seek=position)

        threads = []
        for i in range(3):
            t = threading.Thread(target=writer, args=(i,))
            threads.append(t)
            t.start()

        for t in threads:
            t.join()

        b.record()


@pytest.mark.threading
class TestThreadingDBRecorder:
    """Test cases for DBRecorder thread safety."""

    def test_concurrent_database_writes(self, temp_db):
        """Test multiple threads writing to database."""
        d = DBRecorder(temp_db, cache_size=50)

        def add_data(thread_id):
            for i in range(10):
                d.add_data({'thread': thread_id, 'index': i, 'data': f't{thread_id}_i{i}'}, table='test')

        threads = []
        for i in range(5):
            t = threading.Thread(target=add_data, args=(i,))
            threads.append(t)
            t.start()

        for t in threads:
            t.join()

        d.record()

        # Verify data was written
        import sqlite3
        conn = sqlite3.connect(temp_db)
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM test")
        count = cursor.fetchone()[0]
        assert count == 50  # 5 threads * 10 rows
        conn.close()

    def test_concurrent_different_tables(self, temp_db):
        """Test multiple threads writing to different tables."""
        d = DBRecorder(temp_db, cache_size=100)

        def table_writer(table_name):
            d2 = DBRecorder(temp_db)
            d2.set.table(table_name)
            for i in range(5):
                d2.add_data({'value': i})
            d2.record()

        threads = []
        for i in range(3):
            t = threading.Thread(target=table_writer, args=(f'table_{i}',))
            threads.append(t)
            t.start()

        for t in threads:
            t.join()

        # Verify all tables were created
        tables = d.tables
        assert len(tables) >= 3


@pytest.mark.threading
@pytest.mark.slow
class TestThreadingStress:
    """Stress tests for threading."""

    def test_many_threads_csv(self, temp_csv):
        """Test many threads writing to CSV."""
        r = Recorder(temp_csv, cache_size=1000)

        def writer(thread_id):
            for i in range(20):
                r.add_data((thread_id, i, 'test'))

        threads = []
        for i in range(10):
            t = threading.Thread(target=writer, args=(i,))
            threads.append(t)
            t.start()

        for t in threads:
            t.join()

        r.record()

        # Verify file exists
        assert Path(temp_csv).exists()

    def test_rapid_add_record_cycles(self, temp_csv):
        """Test rapid add_data and record cycles."""
        r = Recorder(temp_csv, cache_size=5)

        def rapid_worker():
            for _ in range(20):
                r.add_data((1, 2, 3))
                sleep(0.001)  # Small delay

        threads = []
        for _ in range(3):
            t = threading.Thread(target=rapid_worker)
            threads.append(t)
            t.start()

        for t in threads:
            t.join()

        r.record()
        assert Path(temp_csv).exists()
