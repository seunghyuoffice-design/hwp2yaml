"""
HWP 5.x/HWPX 파서 - olefile 기반 직접 구현
라이선스: MIT (pyhwp 대체)
"""

from .models import ExtractResult, BatchResult, RecordHeader
from .reader import HWPReader
from .extractor import TextExtractor, extract_hwp_text
from .batch import BatchProcessor
from .exporter import YAMLExporter
from .utils import (
    is_hwpx,
    extract_hwpx_text,
    convert_table_tags_to_markdown,
    clean_text,
)

__version__ = "0.2.0"
__all__ = [
    "HWPReader",
    "TextExtractor",
    "extract_hwp_text",
    "BatchProcessor",
    "YAMLExporter",
    "ExtractResult",
    "BatchResult",
    "RecordHeader",
    "is_hwpx",
    "extract_hwpx_text",
    "convert_table_tags_to_markdown",
    "clean_text",
]
