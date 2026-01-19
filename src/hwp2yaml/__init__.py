"""
hwp2yaml - HWP 5.x / HWPX / HWP 3.x → YAML 변환기
라이선스: MIT (pyhwp 대체)
"""

from .models import ExtractResult, BatchResult, RecordHeader, StructuredResult
from .reader import HWPReader
from .extractor import TextExtractor, extract_hwp_text, extract_hwp_structure
from .batch import BatchProcessor
from .exporter import YAMLExporter
from .utils import (
    is_hwpx,
    extract_hwpx_text,
    extract_hwpx_structure,
    hwpx_structure_to_flat_text,
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
from .structure import (
    DocumentStructure,
    Section,
    Paragraph,
    Table,
    TableCell,
    StructureParser,
    extract_hwp5_structure,
)

__version__ = "0.6.0"
__all__ = [
    # Core
    "HWPReader",
    "TextExtractor",
    "extract_hwp_text",
    "extract_hwp_structure",
    "BatchProcessor",
    "YAMLExporter",
    # Models
    "ExtractResult",
    "BatchResult",
    "RecordHeader",
    "StructuredResult",
    # Structure (HWP 5.x)
    "DocumentStructure",
    "Section",
    "Paragraph",
    "Table",
    "TableCell",
    "StructureParser",
    "extract_hwp5_structure",
    # Utils
    "is_hwpx",
    "extract_hwpx_text",
    "extract_hwpx_structure",
    "hwpx_structure_to_flat_text",
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
