"""HWP 구조 보존 파서 테스트"""

import pytest
from hwp2yaml.structure import (
    StructureParser,
    DocumentStructure,
    Section,
    Paragraph,
    Table,
    TableCell,
    extract_hwp5_structure,
)
from hwp2yaml.models import RecordHeader
from hwp2yaml.record import Record
from hwp2yaml.constants import (
    HWPTAG_PARA_HEADER,
    HWPTAG_PARA_TEXT,
    HWPTAG_CTRL_HEADER,
    HWPTAG_TABLE,
    HWPTAG_LIST_HEADER,
    DEFAULT_ENCODING,
)


def make_record_bytes(tag_id: int, level: int, data: bytes) -> bytes:
    """테스트용 레코드 바이트 생성

    레코드 헤더 구조 (4바이트):
    - bits 0-9: TagID (10비트)
    - bits 10-19: Level (10비트)
    - bits 20-31: Size (12비트)
    """
    size = len(data)
    if size > 0xFFE:
        # 확장 크기 필요
        header_val = tag_id | (level << 10) | (0xFFF << 20)
        return header_val.to_bytes(4, "little") + size.to_bytes(4, "little") + data
    else:
        header_val = tag_id | (level << 10) | (size << 20)
        return header_val.to_bytes(4, "little") + data


def make_para_header(level: int, ctrl_mask: int = 0) -> bytes:
    """단락 헤더 레코드 생성"""
    # PARA_HEADER 구조: nChars(4) + nControlMask(4) + ...
    data = b"\x00\x00\x00\x00"  # nChars
    data += ctrl_mask.to_bytes(4, "little")  # nControlMask
    data += b"\x00" * 16  # 패딩
    return make_record_bytes(HWPTAG_PARA_HEADER, level, data)


def make_para_text(level: int, text: str) -> bytes:
    """단락 텍스트 레코드 생성"""
    encoded = text.encode(DEFAULT_ENCODING)
    return make_record_bytes(HWPTAG_PARA_TEXT, level, encoded)


def make_ctrl_header(level: int, ctrl_id: bytes = b"tbl ") -> bytes:
    """컨트롤 헤더 레코드 생성 (테이블)"""
    return make_record_bytes(HWPTAG_CTRL_HEADER, level, ctrl_id)


def make_table(level: int, rows: int, cols: int) -> bytes:
    """테이블 정의 레코드 생성"""
    # TABLE 구조: 속성(4) + nRows(2) + nCols(2) + ...
    data = b"\x00\x00\x00\x00"  # 속성
    data += rows.to_bytes(2, "little")
    data += cols.to_bytes(2, "little")
    data += b"\x00" * 8  # 패딩
    return make_record_bytes(HWPTAG_TABLE, level, data)


def make_list_header(level: int) -> bytes:
    """리스트(셀) 헤더 레코드 생성"""
    return make_record_bytes(HWPTAG_LIST_HEADER, level, b"\x00" * 8)


class TestTableEndDetection:
    """Fix #1: 테이블 종료 감지 테스트"""

    def test_table_ends_when_level_drops(self):
        """테이블은 레벨이 시작 레벨 아래로 떨어지면 종료"""
        # 구조:
        # Level 0: PARA_HEADER (일반 단락)
        # Level 0: PARA_TEXT "Before table"
        # Level 1: CTRL_HEADER (테이블 시작)
        # Level 2: TABLE (2x2)
        # Level 2: LIST_HEADER (셀 1)
        # Level 2: PARA_HEADER
        # Level 2: PARA_TEXT "Cell1"
        # Level 2: LIST_HEADER (셀 2)
        # Level 2: PARA_HEADER
        # Level 2: PARA_TEXT "Cell2"
        # Level 0: PARA_HEADER (테이블 후 - 레벨 드롭으로 테이블 종료)
        # Level 0: PARA_TEXT "After table"

        section_data = b""
        section_data += make_para_header(0)
        section_data += make_para_text(0, "Before table")
        section_data += make_ctrl_header(1)  # 테이블 시작 레벨 1
        section_data += make_table(2, 1, 2)
        section_data += make_list_header(2)
        section_data += make_para_header(2)
        section_data += make_para_text(2, "Cell1")
        section_data += make_list_header(2)
        section_data += make_para_header(2)
        section_data += make_para_text(2, "Cell2")
        section_data += make_para_header(0)  # 레벨 0으로 드롭 → 테이블 종료
        section_data += make_para_text(0, "After table")

        parser = StructureParser()
        section = parser.parse_section(0, section_data)

        # 테이블 1개
        assert len(section.tables) == 1
        assert section.tables[0].rows == 1
        assert section.tables[0].cols == 2

        # 테이블 후 단락이 별도로 존재 (테이블 셀에 포함되지 않음)
        after_table_texts = [p.text for p in section.paragraphs if "After" in p.text]
        assert len(after_table_texts) == 1
        assert after_table_texts[0] == "After table"


class TestMultiRecordParagraph:
    """Fix #2: 다중 레코드 단락 누적 테스트"""

    def test_accumulates_multiple_para_text(self):
        """하나의 단락에 여러 PARA_TEXT 레코드가 연속될 때 누적"""
        # 일부 HWP 파일에서 긴 단락은 여러 PARA_TEXT로 분할됨
        section_data = b""
        section_data += make_para_header(0)
        section_data += make_para_text(0, "First part ")
        section_data += make_para_text(0, "Second part ")
        section_data += make_para_text(0, "Third part")

        parser = StructureParser()
        section = parser.parse_section(0, section_data)

        # 단락 1개 (텍스트 합쳐짐)
        assert len(section.paragraphs) == 1
        assert "First part" in section.paragraphs[0].text
        assert "Second part" in section.paragraphs[0].text
        assert "Third part" in section.paragraphs[0].text


class TestCellAdvancement:
    """Fix #3: 셀 진행 로직 테스트"""

    def test_cell_advancement_only_in_table(self):
        """LIST_HEADER는 테이블 내부에서만 셀 이동"""
        section_data = b""
        section_data += make_para_header(0)
        section_data += make_para_text(0, "Normal para")
        # 테이블 외부의 LIST_HEADER (셀 이동 안함)
        section_data += make_list_header(0)
        section_data += make_para_header(0)
        section_data += make_para_text(0, "After list header")

        parser = StructureParser()
        section = parser.parse_section(0, section_data)

        # 테이블 없음
        assert len(section.tables) == 0
        # 단락 2개 존재
        assert len(section.paragraphs) >= 2

    def test_cell_bounds_check(self):
        """셀 범위 초과 시 조용히 처리 (크래시 방지)"""
        # 작은 테이블에 많은 셀 데이터
        section_data = b""
        section_data += make_ctrl_header(1)  # 테이블 시작 레벨 1
        section_data += make_table(2, 1, 1)  # 1x1 테이블 (레벨 2)
        section_data += make_list_header(2)
        section_data += make_para_header(2)
        section_data += make_para_text(2, "Cell00")
        # 두 번째 셀 (범위 초과)
        section_data += make_list_header(2)
        section_data += make_para_header(2)
        section_data += make_para_text(2, "Overflow")

        parser = StructureParser()
        section = parser.parse_section(0, section_data)

        # 크래시 없이 파싱 완료
        assert len(section.tables) == 1
        # 첫 번째 셀만 저장됨
        assert len(section.tables[0].cells) >= 1


class TestDocumentStructure:
    """DocumentStructure 모델 테스트"""

    def test_to_flat_text(self):
        """구조를 평탄 텍스트로 변환"""
        doc = DocumentStructure()
        section = Section(index=0)
        section.paragraphs.append(Paragraph(text="Line 1"))
        section.paragraphs.append(Paragraph(text="Line 2"))
        doc.sections.append(section)

        flat = doc.to_flat_text()
        assert "Line 1" in flat
        assert "Line 2" in flat

    def test_get_all_tables(self):
        """모든 테이블 추출"""
        doc = DocumentStructure()
        section1 = Section(index=0)
        section1.tables.append(Table(rows=2, cols=2))
        section2 = Section(index=1)
        section2.tables.append(Table(rows=3, cols=3))
        doc.sections.extend([section1, section2])

        tables = doc.get_all_tables()
        assert len(tables) == 2


class TestTableToDict:
    """테이블 딕셔너리 변환 테스트"""

    def test_table_to_dict(self):
        """테이블을 2D 배열 딕셔너리로 변환"""
        table = Table(rows=2, cols=2)
        table.cells.append(TableCell(row=0, col=0, text="A"))
        table.cells.append(TableCell(row=0, col=1, text="B"))
        table.cells.append(TableCell(row=1, col=0, text="C"))
        table.cells.append(TableCell(row=1, col=1, text="D"))

        d = table.to_dict()
        assert d["rows"] == 2
        assert d["cols"] == 2
        assert d["data"][0][0] == "A"
        assert d["data"][1][1] == "D"
