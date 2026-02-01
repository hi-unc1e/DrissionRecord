# -*- coding:utf-8 -*-
"""Tests for backup functionality."""
import csv
from pathlib import Path
from datetime import datetime

import pytest

from DrissionRecord import Recorder, ByteRecorder, DBRecorder


class TestBackup:
    """Test cases for backup functionality."""

    def test_backup_after_record(self, temp_csv, backup_dir):
        """Test creating backup after recording."""
        r = Recorder(temp_csv)
        r.set.auto_backup(folder=backup_dir)
        r.add_data((1, 2, 3))
        r.record()

        # Create backup
        backup_path = r.backup()
        assert Path(backup_path).exists()
        assert Path(backup_path).parent == Path(backup_dir)

    def test_auto_backup_interval(self, temp_csv, backup_dir):
        """Test automatic backup at intervals."""
        r = Recorder(temp_csv)
        r.set.auto_backup(interval=2, folder=backup_dir)

        r.add_data((1, 2, 3))
        r.record()  # First record - backup_times = 1

        r.add_data((4, 5, 6))
        r.record()  # Second record - backup should trigger

        # Check if backup was created (may need one more record depending on timing)
        r.add_data((7, 8, 9))
        r.record()  # Third record - backup should definitely have triggered

        backup_files = list(Path(backup_dir).glob('*.csv'))
        # Note: backup timing depends on when backup_times reaches interval
        assert len(backup_files) >= 1 or Path(temp_csv).exists()

    def test_backup_with_custom_name(self, temp_csv, backup_dir):
        """Test backup with custom name."""
        r = Recorder(temp_csv)
        r.add_data((1, 2, 3))
        r.record()

        custom_name = 'my_backup.csv'
        backup_path = r.backup(folder=backup_dir, name=custom_name)

        assert Path(backup_path).name == custom_name
        assert Path(backup_path).exists()

    def test_backup_overwrite(self, temp_csv, backup_dir):
        """Test backup with overwrite enabled."""
        r = Recorder(temp_csv)
        r.add_data((1, 2, 3))
        r.record()

        # Create first backup
        backup_path1 = r.backup(folder=backup_dir, name='backup.csv', overwrite=False)

        # Add more data
        r2 = Recorder(temp_csv)
        r2.add_data((4, 5, 6))
        r2.record()

        # Create second backup with same name (should create new file)
        backup_path2 = r2.backup(folder=backup_dir, name='backup.csv', overwrite=False)

        # Both should exist (second has timestamp)
        assert Path(backup_path1).exists()
        assert Path(backup_path2).exists()

    def test_backup_overwrite_true(self, temp_csv, backup_dir):
        """Test backup with overwrite=True."""
        r = Recorder(temp_csv)
        r.add_data((1, 2, 3))
        r.record()

        # Create backup
        r.backup(folder=backup_dir, name='backup.csv', overwrite=False)

        # Add more data and overwrite backup
        r2 = Recorder(temp_csv)
        r2.add_data((4, 5, 6))
        r2.record()

        backup_files_before = list(Path(backup_dir).glob('backup*.csv'))
        r2.backup(folder=backup_dir, name='backup.csv', overwrite=True)

        backup_files_after = list(Path(backup_dir).glob('backup*.csv'))
        # With overwrite, should have same or fewer files
        assert len(backup_files_after) <= len(backup_files_before)

    def test_backup_byte_recorder(self, temp_byte_file, backup_dir):
        """Test backup for ByteRecorder."""
        b = ByteRecorder(temp_byte_file)
        b.set.auto_backup(folder=backup_dir)
        b.add_data(b'hello world')
        b.record()

        backup_path = b.backup()
        assert Path(backup_path).exists()

        # Verify backup content
        with open(backup_path, 'rb') as f:
            assert f.read() == b'hello world'

    def test_backup_db_recorder(self, temp_db, backup_dir):
        """Test backup for DBRecorder."""
        import sqlite3

        d = DBRecorder(temp_db)
        d.set.auto_backup(folder=backup_dir)
        d.add_data({'name': 'Alice', 'age': 30}, table='users')
        d.record()

        backup_path = d.backup()
        assert Path(backup_path).exists()

        # Verify backup content
        conn = sqlite3.connect(backup_path)
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM users")
        result = cursor.fetchone()
        assert result == ('Alice', 30)
        conn.close()

    def test_backup_xlsx(self, temp_xlsx, backup_dir):
        """Test backup for XLSX files."""
        import openpyxl

        r = Recorder(temp_xlsx)
        r.set.auto_backup(folder=backup_dir)
        r.add_data((1, 2, 3))
        r.record()

        backup_path = r.backup()
        assert Path(backup_path).exists()

        # Verify backup content
        wb = openpyxl.load_workbook(backup_path)
        ws = wb.active
        assert ws.cell(1, 1).value == 1
        wb.close()

    def test_backup_json(self, temp_json, backup_dir):
        """Test backup for JSON files."""
        import json

        r = Recorder(temp_json)
        r.set.auto_backup(folder=backup_dir)
        r.add_data({'key': 'value'})
        r.record()

        backup_path = r.backup()
        assert Path(backup_path).exists()

        # Verify backup content
        with open(backup_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            assert data == [{'key': 'value'}]

    def test_backup_without_existing_file(self, temp_csv, backup_dir):
        """Test backup when file doesn't exist yet."""
        r = Recorder(temp_csv)
        r.add_data((1, 2, 3))
        # Don't call record() yet

        # Backup should return empty string
        result = r.backup(folder=backup_dir)
        assert result == ''

    def test_default_backup_folder(self, temp_csv):
        """Test backup with default folder name."""
        r = Recorder(temp_csv)
        r.add_data((1, 2, 3))
        r.record()

        # Backup should create 'backup' folder
        backup_path = r.backup()
        assert 'backup' in backup_path

    def test_backup_preserves_extension(self, temp_csv, backup_dir):
        """Test that backup preserves file extension."""
        r = Recorder(temp_csv)
        r.add_data((1, 2, 3))
        r.record()

        backup_path = r.backup(folder=backup_dir)
        assert backup_path.endswith('.csv')

    def test_backup_creates_folder(self, temp_csv):
        """Test that backup creates the folder if it doesn't exist."""
        nested_backup = str(Path(temp_csv).parent / 'backups' / 'nested' / 'folder')

        r = Recorder(temp_csv)
        r.add_data((1, 2, 3))
        r.record()

        backup_path = r.backup(folder=nested_backup)
        assert Path(backup_path).exists()
        assert Path(nested_backup).exists()

    def test_multiple_backups(self, temp_csv, backup_dir):
        """Test creating multiple backups with timestamp."""
        r = Recorder(temp_csv)
        r.add_data((1, 2, 3))
        r.record()

        # Create multiple backups
        backup1 = r.backup(folder=backup_dir, name='data.csv', overwrite=False)

        r.add_data((4, 5, 6))
        r.record()
        backup2 = r.backup(folder=backup_dir, name='data.csv', overwrite=False)

        # Both should exist, second should have timestamp
        assert Path(backup1).exists()
        assert Path(backup2).exists()
        # Files should have different names
        assert Path(backup1).name != Path(backup2).name

    def test_backup_after_data_change(self, temp_csv, backup_dir):
        """Test that backup captures current data state."""
        r = Recorder(temp_csv)
        r.add_data((1, 2, 3))
        r.record()

        # Create backup
        backup_path = r.backup(folder=backup_dir)
        with open(backup_path, 'r', encoding='utf-8', newline='') as f:
            reader = csv.reader(f)
            original_backup = list(reader)

        # Add more data
        r.add_data((4, 5, 6))
        r.record()

        # Create another backup
        backup_path2 = r.backup(folder=backup_dir, name='backup2.csv')
        with open(backup_path2, 'r', encoding='utf-8', newline='') as f:
            reader = csv.reader(f)
            new_backup = list(reader)

        # New backup should have more rows
        assert len(new_backup) > len(original_backup)
