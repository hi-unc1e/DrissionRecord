# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

DrissionRecord is a Python library for recording data to files with efficient caching, multi-threading support, and automatic backup. Supports csv, xlsx, json, jsonl, txt files and SQLite databases.

**Core Philosophy:** Simple, reliable, worry-free.

## Build and Installation

```bash
# Install the package in development mode
pip install -e .

# Or using setup.py directly
python setup.py install

# Install dependencies
pip install -r requirements.txt
```

**Dependencies:**
- `openpyxl` - For Excel (xlsx) file support
- Python 3.6+

## Architecture

### Class Hierarchy

```
OriginalRecorder (abstract base)
    ├── BaseRecorder (adds table/encoding support)
    │   ├── Recorder (csv, xlsx, json, jsonl, txt)
    │   └── DBRecorder (SQLite)
    └── ByteRecorder (binary data)
```

### Key Components

- **DrissionRecord/base.py** - `OriginalRecorder` (thread-safe base with caching), `BaseRecorder` (adds multi-table support)
- **DrissionRecord/recorder.py** - `Recorder` class for file formats
- **DrissionRecord/byte_recorder.py** - `ByteRecorder` for binary data
- **DrissionRecord/db_recorder.py** - `DBRecorder` for SQLite
- **DrissionRecord/cell_style.py** - `CellStyle` for Excel formatting
- **DrissionRecord/setter.py** - Setter classes for fluent API (`.set` property)
- **DrissionRecord/tools.py** - Utilities including `Col` class

### Design Patterns

1. **Fluent API**: All recorders have a `.set` property for method-chaining configuration
2. **Thread Safety**: Uses `threading.Lock` for concurrent operations
3. **Caching**: Data is cached in memory and batch-written for performance
4. **Auto-recovery**: Automatic data saving on exceptions and object destruction
5. **Coordinate System**: Supports multiple formats: `(row, col)`, `'A1'`, `(row, 'header')`

### Data Flow

1. Create recorder object with file path
2. Optionally configure using `.set` property
3. Call `add_data()` to add data (cached in memory)
4. Data automatically written when cache is full or explicitly via `record()`
5. On destruction, any remaining cached data is automatically saved

### Header Management

Headers are stored as `{column_index: header_value}` dictionaries. For xlsx files with multiple sheets:
```python
{None: Header, 'sheet1': Header1, 'sheet2': Header2, ...}
```
`None` represents the active sheet in xlsx, or the only sheet in csv.

## Coordinate Systems

The library supports multiple coordinate formats for data insertion:
- Tuple: `(row, col)` - zero-based row, col indices
- A1 notation: `'A1'`, `'B3'` etc.
- Header-based: `(row, 'header_name')` - uses column with matching header

## Common Development Tasks

### Adding a new file format support

1. Extend `BaseRecorder` in `recorder.py`
2. Implement `_read_file()` and `_write_data()` methods
3. Add format detection logic
4. Export from `__init__.py`

### Adding new features to existing recorders

- Modify the specific recorder class (`recorder.py`, `byte_recorder.py`, or `db_recorder.py`)
- For cross-cutting concerns, add to base classes in `base.py`
- Update corresponding `.pyi` stub file for type hints

## Testing

This project currently has no test suite. When adding tests:
- Create a `tests/` directory
- Use `pytest` as the test framework
- Test multi-threading scenarios thoroughly
- Test auto-recovery behavior (simulated crashes)
