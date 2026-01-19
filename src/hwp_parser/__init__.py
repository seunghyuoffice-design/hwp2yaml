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
from .triage import (
    HWPVersion,
    TriageResult,
    TriageSummary,
    detect_hwp_version,
    triage_file,
    triage_files,
    triage_directory,
)
from .hwp3_converter import (
    HWP3Converter,
    ConversionResult,
    convert_hwp3,
    batch_convert_hwp3,
    convert_to_yaml,
    batch_convert_to_yaml,
)

__version__ = "0.5.0"
__all__ = [
    # Core
    "HWPReader",
    "TextExtractor",
    "extract_hwp_text",
    "BatchProcessor",
    "YAMLExporter",
    # Models
    "ExtractResult",
    "BatchResult",
    "RecordHeader",
    # Utils
    "is_hwpx",
    "extract_hwpx_text",
    "convert_table_tags_to_markdown",
    "clean_text",
    # Triage
    "HWPVersion",
    "TriageResult",
    "TriageSummary",
    "detect_hwp_version",
    "triage_file",
    "triage_files",
    "triage_directory",
    # HWP 3.x Converter
    "HWP3Converter",
    "ConversionResult",
    "convert_hwp3",
    "batch_convert_hwp3",
    "convert_to_yaml",
    "batch_convert_to_yaml",
]
