# -*- coding:utf-8 -*-
from pathlib import Path
from typing import Union, Optional

from .base import OriginalRecorder


class ByteRecorder(OriginalRecorder):
    _SUPPORTS: tuple = ...
    __END: tuple = ...
    data: list = ...

    def __init__(self,
                 path: Union[None, str, Path] = None,
                 cache_size: int = 1000): ...

    def add_data(self,
                 data: bytes,
                 seek: int = None) -> None: ...

    def _record(self) -> None: ...
