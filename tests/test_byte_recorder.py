# -*- coding:utf-8 -*-
"""Tests for ByteRecorder class."""
from pathlib import Path

import pytest

from DrissionRecord import ByteRecorder


class TestByteRecorder:
    """Test cases for ByteRecorder class."""

    def test_create_byte_recorder(self, temp_byte_file):
        """Test creating a ByteRecorder."""
        b = ByteRecorder(temp_byte_file)
        assert b.path == temp_byte_file
        assert b.type == 'byte'

    def test_add_data_bytes(self, temp_byte_file):
        """Test adding bytes data."""
        b = ByteRecorder(temp_byte_file)
        b.add_data(b'hello world')
        b.record()

        with open(temp_byte_file, 'rb') as f:
            content = f.read()
            assert content == b'hello world'

    def test_add_multiple_data(self, temp_byte_file):
        """Test adding multiple byte chunks."""
        b = ByteRecorder(temp_byte_file)
        b.add_data(b'hello ')
        b.add_data(b'world ')
        b.add_data(b'!')
        b.record()

        with open(temp_byte_file, 'rb') as f:
            content = f.read()
            assert content == b'hello world !'

    def test_cache_size_auto_record(self, temp_byte_file):
        """Test automatic recording when cache size is reached."""
        b = ByteRecorder(temp_byte_file, cache_size=2)
        b.add_data(b'first')
        b.add_data(b'second')
        # Should auto-record after cache_size is reached

        # File should exist now
        assert Path(temp_byte_file).exists()

        # Third addition should also trigger record
        b.add_data(b'third')
        assert Path(temp_byte_file).exists()

    def test_add_data_with_seek(self, temp_byte_file):
        """Test adding data with seek position."""
        b = ByteRecorder(temp_byte_file)
        b.add_data(b'hello world')
        b.record()

        # Add more data at position 5
        b2 = ByteRecorder(temp_byte_file)
        b2.add_data(b'XXXXX', seek=5)
        b2.record()

        with open(temp_byte_file, 'rb') as f:
            content = f.read()
            # Should overwrite position 5-10
            assert content[:5] == b'hello'
            assert content[5:10] == b'XXXXX'

    def test_seek_at_beginning(self, temp_byte_file):
        """Test seeking to the beginning."""
        b = ByteRecorder(temp_byte_file)
        b.add_data(b'original data')
        b.record()

        # Overwrite from beginning
        b2 = ByteRecorder(temp_byte_file)
        b2.add_data(b'NEW', seek=0)
        b2.record()

        with open(temp_byte_file, 'rb') as f:
            content = f.read()
            assert content[:3] == b'NEW'

    def test_type_error_non_bytes(self, temp_byte_file):
        """Test that non-bytes data raises TypeError."""
        b = ByteRecorder(temp_byte_file)
        with pytest.raises(TypeError, match='只能接受bytes类型数据'):
            b.add_data('string not bytes')

    def test_invalid_seek_value(self, temp_byte_file):
        """Test that invalid seek values raise ValueError."""
        b = ByteRecorder(temp_byte_file)
        with pytest.raises(ValueError, match='seek参数只能接受None或大于等于0的整数'):
            b.add_data(b'data', seek=-1)

        with pytest.raises(ValueError, match='seek参数只能接受None或大于等于0的整数'):
            b.add_data(b'data', seek='invalid')

    def test_clear_data(self, temp_byte_file):
        """Test clearing cached data."""
        b = ByteRecorder(temp_byte_file, cache_size=100)
        b.add_data(b'some data')
        b.clear()
        assert b._data_count == 0
        assert len(b.data) == 0

    def test_delete_method(self, temp_byte_file):
        """Test the delete() method."""
        b = ByteRecorder(temp_byte_file)
        b.add_data(b'test data')
        b.record()
        assert Path(temp_byte_file).exists()

        b.delete()
        assert not Path(temp_byte_file).exists()

    def test_empty_data(self, temp_byte_file):
        """Test handling empty byte data."""
        b = ByteRecorder(temp_byte_file)
        b.add_data(b'')
        b.record()

        # File should be created
        assert Path(temp_byte_file).exists()

    def test_large_binary_data(self, temp_byte_file):
        """Test handling larger binary data."""
        b = ByteRecorder(temp_byte_file)
        large_data = b'x' * 10000
        b.add_data(large_data)
        b.record()

        with open(temp_byte_file, 'rb') as f:
            content = f.read()
            assert len(content) == 10000
            assert content == large_data

    def test_append_to_existing_file(self, temp_byte_file):
        """Test appending to existing binary file."""
        b = ByteRecorder(temp_byte_file)
        b.add_data(b'first ')
        b.record()

        b2 = ByteRecorder(temp_byte_file)
        b2.add_data(b'second')
        b2.record()

        with open(temp_byte_file, 'rb') as f:
            content = f.read()
            assert content == b'first second'

    def test_binary_zero_bytes(self, temp_byte_file):
        """Test handling zero bytes in binary data."""
        b = ByteRecorder(temp_byte_file)
        b.add_data(b'\x00\x01\x02\xff\xfe')
        b.record()

        with open(temp_byte_file, 'rb') as f:
            content = f.read()
            assert content == b'\x00\x01\x02\xff\xfe'

    def test_unicode_bytes(self, temp_byte_file):
        """Test handling Unicode characters as bytes."""
        b = ByteRecorder(temp_byte_file)
        b.add_data('你好世界'.encode('utf-8'))
        b.record()

        with open(temp_byte_file, 'rb') as f:
            content = f.read()
            assert content == '你好世界'.encode('utf-8')

    def test_cache_size_zero(self, temp_byte_file):
        """Test with cache_size=0 (no auto-record)."""
        b = ByteRecorder(temp_byte_file, cache_size=0)
        b.add_data(b'test')
        # Should not auto-record

        # Explicitly record
        b.record()
        assert Path(temp_byte_file).exists()

    def test_show_msg_false(self, temp_byte_file, capsys):
        """Test disabling messages."""
        b = ByteRecorder(temp_byte_file)
        b.set.show_msg(False)
        b.add_data(b'test')
        b.record()

        captured = capsys.readouterr()
        assert '开始写入文件' not in captured.out

    def test_set_path_using_setter(self, temp_dir):
        """Test setting path using the fluent API."""
        b = ByteRecorder()
        b.set.path(str(Path(temp_dir) / 'test.bin'))
        assert b.path == str(Path(temp_dir) / 'test.bin')
