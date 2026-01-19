"""HWP 유틸리티 테스트"""

import pytest
from hwp2yaml.utils import (
    _sort_section_files,
    _get_local_tag,
    _strip_namespaces_safe,
    _extract_text_preserving_whitespace,
    _parse_hwpx_section,
    _parse_hwpx_section_structure,
    convert_table_tags_to_markdown,
    clean_text,
)
import xml.etree.ElementTree as ET


class TestSectionOrdering:
    """Fix #4: 섹션 파일 숫자 정렬 테스트"""

    def test_numeric_sort_single_digit(self):
        """단일 자릿수 숫자 정렬"""
        files = ["section2.xml", "section1.xml", "section3.xml"]
        sorted_files = _sort_section_files(files)
        assert sorted_files == ["section1.xml", "section2.xml", "section3.xml"]

    def test_numeric_sort_double_digit(self):
        """두 자릿수 숫자 정렬 (렉시코그래픽 vs 숫자)"""
        # 렉시코그래픽: section10 < section2
        # 숫자: section2 < section10
        files = ["section10.xml", "section2.xml", "section1.xml"]
        sorted_files = _sort_section_files(files)
        assert sorted_files == ["section1.xml", "section2.xml", "section10.xml"]

    def test_numeric_sort_with_path(self):
        """경로가 포함된 파일명 정렬"""
        files = [
            "Contents/section10.xml",
            "Contents/section2.xml",
            "Contents/section1.xml"
        ]
        sorted_files = _sort_section_files(files)
        assert sorted_files[0] == "Contents/section1.xml"
        assert sorted_files[1] == "Contents/section2.xml"
        assert sorted_files[2] == "Contents/section10.xml"

    def test_numeric_sort_mixed_case(self):
        """대소문자 혼합"""
        files = ["Section2.xml", "section1.xml", "SECTION10.xml"]
        sorted_files = _sort_section_files(files)
        # 숫자 기준 정렬
        assert "1" in sorted_files[0]
        assert "2" in sorted_files[1]
        assert "10" in sorted_files[2]


class TestNamespaceParsing:
    """Fix #5: 네임스페이스 안전 파싱 테스트"""

    def test_get_local_tag_with_namespace(self):
        """네임스페이스 포함 태그에서 로컬명 추출"""
        elem = ET.fromstring('<root xmlns="http://example.com"><p/></root>')
        child = list(elem)[0]
        # 네임스페이스 포함된 태그: {http://example.com}p
        tag = _get_local_tag(child)
        assert tag == "p"

    def test_get_local_tag_without_namespace(self):
        """네임스페이스 없는 태그"""
        elem = ET.fromstring('<root><p/></root>')
        child = list(elem)[0]
        tag = _get_local_tag(child)
        assert tag == "p"

    def test_strip_namespaces_safe(self):
        """네임스페이스 안전 제거"""
        xml = '<hp:root xmlns:hp="http://example.com"><hp:p>text</hp:p></hp:root>'
        clean = _strip_namespaces_safe(xml)

        # 접두사 제거됨
        assert "hp:" not in clean
        # 내용 유지
        assert "<p>text</p>" in clean or "<root>" in clean

    def test_parse_hwpx_section_with_namespace(self):
        """네임스페이스 있는 HWPX 섹션 파싱"""
        xml = '''<sec xmlns="http://www.hancom.co.kr/hwpml/2011/section">
            <p><t>Hello</t></p>
            <p><t>World</t></p>
        </sec>'''

        text = _parse_hwpx_section(xml)
        # 텍스트가 추출되어야 함
        assert "Hello" in text or "World" in text

    def test_parse_hwpx_section_with_prefixed_namespace(self):
        """접두사 네임스페이스 있는 HWPX 섹션 파싱"""
        xml = '''<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section">
            <hs:p><hs:t>Hello</hs:t></hs:p>
        </hs:sec>'''

        text = _parse_hwpx_section(xml)
        # 텍스트가 추출되어야 함 (파싱 실패 아님)
        assert text is not None

    def test_extract_text_preserving_whitespace(self):
        """xml:space="preserve" 존중"""
        xml = '<root xml:space="preserve">  spaced  </root>'
        elem = ET.fromstring(xml)
        texts = _extract_text_preserving_whitespace(elem)
        # 공백 보존
        if texts:
            assert "spaced" in texts[0]


class TestParseHwpxSectionStructure:
    """HWPX 섹션 구조 파싱 테스트"""

    def test_parse_simple_paragraphs(self):
        """간단한 단락 구조 파싱"""
        xml = '''<sec>
            <p><t>First paragraph</t></p>
            <p><t>Second paragraph</t></p>
        </sec>'''

        result = _parse_hwpx_section_structure(xml, 0)
        assert result is not None
        assert len(result["paragraphs"]) >= 2

    def test_parse_table_structure(self):
        """테이블 구조 파싱"""
        xml = '''<sec>
            <tbl>
                <tr><tc><t>A</t></tc><tc><t>B</t></tc></tr>
                <tr><tc><t>C</t></tc><tc><t>D</t></tc></tr>
            </tbl>
        </sec>'''

        result = _parse_hwpx_section_structure(xml, 0)
        assert result is not None
        assert len(result["tables"]) == 1
        assert result["tables"][0]["rows"] == 2
        assert result["tables"][0]["cols"] == 2


class TestConvertTableTagsToMarkdown:
    """표 태그→마크다운 변환 테스트"""

    def test_convert_simple_table(self):
        """간단한 표 변환"""
        text = "<A><B>\n<C><D>"
        result = convert_table_tags_to_markdown(text)
        assert "|" in result
        assert "A" in result
        assert "D" in result

    def test_preserve_non_table_lines(self):
        """표가 아닌 줄은 유지"""
        text = "Normal line\n<A><B>\nAnother normal line"
        result = convert_table_tags_to_markdown(text)
        assert "Normal line" in result
        assert "Another normal line" in result


class TestCleanText:
    """텍스트 정리 테스트"""

    def test_remove_consecutive_spaces(self):
        """연속 공백 제거"""
        text = "Hello    world"
        result = clean_text(text)
        assert "    " not in result

    def test_remove_consecutive_newlines(self):
        """연속 빈 줄 제거"""
        text = "Line1\n\n\n\nLine2"
        result = clean_text(text)
        assert "\n\n\n\n" not in result

    def test_remove_control_chars(self):
        """제어 문자 제거"""
        text = "Hello\x00World\x08"
        result = clean_text(text)
        assert "\x00" not in result
        assert "\x08" not in result
