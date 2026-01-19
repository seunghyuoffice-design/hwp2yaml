"""HWP 추출기 테스트"""

import pytest
from hwp2yaml.extractor import extract_hwp_text
from hwp2yaml.models import ExtractResult


def test_extract_nonexistent_file():
    """존재하지 않는 파일 처리"""
    result = extract_hwp_text("/nonexistent/file.hwp")
    assert not result.success
    assert result.method == "failed"
    assert result.error is not None


def test_extract_result_dataclass():
    """ExtractResult 데이터클래스 테스트"""
    result = ExtractResult(
        filepath="/test.hwp",
        success=True,
        text="테스트 텍스트",
        method="prvtext",
    )
    assert result.char_count == 7
    assert result.filepath == "/test.hwp"
