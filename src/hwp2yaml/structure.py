"""HWP 5.x 구조 보존 파서

문서 구조(단락, 테이블, 섹션)를 보존하여 YAML로 변환
"""

from dataclasses import dataclass, field
from typing import Iterator, Any
from enum import Enum, auto

from .constants import (
    HWPTAG_PARA_HEADER,
    HWPTAG_PARA_TEXT,
    HWPTAG_CTRL_HEADER,
    HWPTAG_TABLE,
    HWPTAG_LIST_HEADER,
    HWPTAG_SHAPE_COMPONENT,
    DEFAULT_ENCODING,
    CTRL_CHAR_PARA_BREAK,
    CTRL_CHAR_LINE_BREAK,
)
from .record import RecordParser, Record


class ControlType(Enum):
    """컨트롤 타입"""
    TABLE = auto()      # tbl
    SHAPE = auto()      # gso
    EQUATION = auto()   # eqed
    FIELD = auto()      # various field types
    UNKNOWN = auto()


@dataclass
class TableCell:
    """테이블 셀"""
    row: int
    col: int
    text: str
    row_span: int = 1
    col_span: int = 1


@dataclass
class Table:
    """테이블 구조"""
    rows: int
    cols: int
    cells: list[TableCell] = field(default_factory=list)

    def to_dict(self) -> dict:
        """딕셔너리 변환"""
        # 2D 배열로 변환
        grid = [["" for _ in range(self.cols)] for _ in range(self.rows)]
        for cell in self.cells:
            if 0 <= cell.row < self.rows and 0 <= cell.col < self.cols:
                grid[cell.row][cell.col] = cell.text

        return {
            "rows": self.rows,
            "cols": self.cols,
            "data": grid,
        }


@dataclass
class Paragraph:
    """단락 구조"""
    text: str
    level: int = 0
    style_id: int = 0

    def to_dict(self) -> dict:
        return {
            "text": self.text,
            "level": self.level,
        }


@dataclass
class Section:
    """섹션 구조"""
    index: int
    paragraphs: list[Paragraph] = field(default_factory=list)
    tables: list[Table] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "index": self.index,
            "paragraphs": [p.to_dict() for p in self.paragraphs],
            "tables": [t.to_dict() for t in self.tables],
        }


@dataclass
class DocumentStructure:
    """전체 문서 구조"""
    sections: list[Section] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "sections": [s.to_dict() for s in self.sections],
        }

    def to_flat_text(self) -> str:
        """평탄화된 텍스트 추출"""
        texts = []
        for section in self.sections:
            for para in section.paragraphs:
                if para.text.strip():
                    texts.append(para.text)
        return "\n".join(texts)

    def get_all_tables(self) -> list[Table]:
        """모든 테이블 추출"""
        tables = []
        for section in self.sections:
            tables.extend(section.tables)
        return tables


class StructureParser:
    """HWP 5.x 구조 파서

    레코드의 level 필드를 활용하여 계층 구조 파악
    """

    # 컨트롤 타입 시그니처 (4바이트, little-endian)
    CTRL_TABLE = b"tbl "      # 0x20 6C 62 74
    CTRL_SHAPE = b"gso "      # 그리기 객체
    CTRL_EQUATION = b"eqed"   # 수식

    def __init__(self):
        self.current_section: Section | None = None
        self.current_table: Table | None = None
        self.current_cell_row: int = 0
        self.current_cell_col: int = 0
        self.in_table: bool = False
        self.in_cell: bool = False  # Fix #3: 셀 내부 상태 추적
        self.cell_texts: list[str] = []
        self.table_start_level: int | None = None  # Fix #1: 테이블 시작 레벨 추적
        self.current_para_texts: list[str] = []  # Fix #2: 다중 레코드 단락 누적

    def parse_section(self, section_idx: int, section_data: bytes) -> Section:
        """섹션 데이터 파싱

        Args:
            section_idx: 섹션 인덱스
            section_data: 압축 해제된 섹션 바이트 데이터

        Returns:
            Section 객체
        """
        section = Section(index=section_idx)
        self.current_section = section
        self.in_table = False
        self.in_cell = False  # Fix #3: 셀 상태 초기화
        self.current_table = None
        self.table_start_level = None  # Fix #1: 테이블 시작 레벨 초기화
        self.current_para_texts = []  # Fix #2: 단락 텍스트 버퍼 초기화

        records = list(RecordParser.iter_records(section_data))

        i = 0
        while i < len(records):
            record = records[i]
            tag_id = record.header.tag_id
            level = record.header.level

            # Fix #1: 테이블 종료 감지 - 레벨이 테이블 시작 레벨 미만으로 떨어지면 종료
            if self.in_table and self.table_start_level is not None:
                if level < self.table_start_level:
                    self._finalize_current_table(section)

            if tag_id == HWPTAG_PARA_HEADER:
                # Fix #2: 이전 단락 텍스트 마무리
                self._finalize_paragraph(section, level)

                # 단락 헤더 - 컨트롤 마스크 확인
                ctrl_mask = self._parse_para_header(record.data)
                self.current_para_texts = []  # 새 단락 시작

            elif tag_id == HWPTAG_PARA_TEXT:
                # Fix #2: 다중 레코드 단락 누적
                text = self._decode_para_text(record.data)
                if text:
                    self.current_para_texts.append(text)

            elif tag_id == HWPTAG_CTRL_HEADER:
                # 컨트롤 헤더 - 타입 확인
                ctrl_type = self._parse_ctrl_header(record.data)

                if ctrl_type == ControlType.TABLE:
                    # Fix #2: 테이블 전 대기 중인 단락 마무리
                    self._finalize_paragraph(section, level)

                    # 이전 테이블이 있으면 먼저 저장
                    if self.in_table and self.current_table:
                        self._finalize_current_table(section)

                    # 새 테이블 시작
                    self.in_table = True
                    self.in_cell = False
                    self.table_start_level = level  # Fix #1: 테이블 시작 레벨 저장
                    self.cell_texts = []
                    self.current_cell_row = 0
                    self.current_cell_col = 0

            elif tag_id == HWPTAG_TABLE:
                # 테이블 정의
                rows, cols = self._parse_table_header(record.data)
                self.current_table = Table(rows=rows, cols=cols)
                self.current_cell_row = 0
                self.current_cell_col = 0

            elif tag_id == HWPTAG_LIST_HEADER:
                # Fix #3: LIST_HEADER 전에 대기 중인 단락 finalize
                if self.in_table:
                    self._finalize_paragraph(section, level)

                # Fix #3: 테이블 내부이고 셀 진입 전에만 셀 이동
                if self.in_table and self.current_table and not self.in_cell:
                    # 이전 셀 텍스트가 있으면 저장
                    if self.cell_texts:
                        self._save_current_cell()

                    # 셀 진입 상태로 변경
                    self.in_cell = True
                elif self.in_table and self.in_cell:
                    # 이미 셀 안에 있으면 현재 셀 저장 후 다음 셀로
                    if self.cell_texts:
                        self._save_current_cell()
                    self._advance_to_next_cell()
                    self.in_cell = True

            i += 1

        # Fix #2: 마지막 단락 처리
        self._finalize_paragraph(section, 0)

        # 마지막 테이블 처리
        if self.in_table and self.current_table:
            self._finalize_current_table(section)

        return section

    def _finalize_paragraph(self, section: Section, level: int) -> None:
        """Fix #2: 누적된 단락 텍스트를 단락 또는 셀에 추가"""
        if not self.current_para_texts:
            return

        full_text = "".join(self.current_para_texts)
        self.current_para_texts = []

        if not full_text.strip():
            return

        if self.in_table:
            # 테이블 셀 내부 텍스트
            self.cell_texts.append(full_text)
        else:
            # 일반 단락
            para = Paragraph(text=full_text, level=level)
            section.paragraphs.append(para)

    def _finalize_current_table(self, section: Section) -> None:
        """Fix #1: 현재 테이블을 마무리하고 섹션에 추가"""
        if not self.current_table:
            return

        # 마지막 셀 저장
        if self.cell_texts:
            self._save_current_cell()

        section.tables.append(self.current_table)
        self.current_table = None
        self.in_table = False
        self.in_cell = False
        self.table_start_level = None
        self.cell_texts = []

    def _save_current_cell(self) -> None:
        """Fix #3: 현재 셀 텍스트를 테이블에 저장 (범위 검사 포함)"""
        if not self.current_table or not self.cell_texts:
            return

        # Fix #3: 범위 검사
        if not (0 <= self.current_cell_row < self.current_table.rows and
                0 <= self.current_cell_col < self.current_table.cols):
            # 범위 초과 시 경고 로그만 남기고 진행 (silent drop 대신)
            self.cell_texts = []
            return

        cell_text = "\n".join(self.cell_texts)
        cell = TableCell(
            row=self.current_cell_row,
            col=self.current_cell_col,
            text=cell_text.strip(),
        )
        self.current_table.cells.append(cell)
        self.cell_texts = []

    def _advance_to_next_cell(self) -> None:
        """다음 셀로 이동"""
        if not self.current_table:
            return

        self.current_cell_col += 1
        if self.current_cell_col >= self.current_table.cols:
            self.current_cell_col = 0
            self.current_cell_row += 1

    def _parse_para_header(self, data: bytes) -> int:
        """PARA_HEADER 파싱

        구조 (최소 22바이트):
        - offset 0-3: nChars (4바이트) - 텍스트 글자 수
        - offset 4-7: nControlMask (4바이트) - 컨트롤 마스크
        - offset 8-11: ParaShapeID (4바이트)
        - offset 12-13: ParaStyleID (2바이트)
        - ...

        Returns:
            컨트롤 마스크 값
        """
        if len(data) < 8:
            return 0

        ctrl_mask = int.from_bytes(data[4:8], "little")
        return ctrl_mask

    def _parse_ctrl_header(self, data: bytes) -> ControlType:
        """CTRL_HEADER 파싱

        구조:
        - offset 0-3: ctrlId (4바이트) - 컨트롤 ID (역순 문자열)

        Returns:
            ControlType enum
        """
        if len(data) < 4:
            return ControlType.UNKNOWN

        ctrl_id = data[0:4]

        # 컨트롤 ID는 리틀 엔디안으로 저장된 4글자 문자열
        # "tbl " -> b"\x20lbt" -> b"tbl "
        if ctrl_id == self.CTRL_TABLE or ctrl_id == self.CTRL_TABLE[::-1]:
            return ControlType.TABLE
        elif ctrl_id == self.CTRL_SHAPE or ctrl_id == self.CTRL_SHAPE[::-1]:
            return ControlType.SHAPE
        elif ctrl_id == self.CTRL_EQUATION or ctrl_id == self.CTRL_EQUATION[::-1]:
            return ControlType.EQUATION

        return ControlType.UNKNOWN

    def _parse_table_header(self, data: bytes) -> tuple[int, int]:
        """TABLE 레코드 파싱

        구조:
        - offset 0-3: 속성 플래그
        - offset 4-5: nRows (2바이트)
        - offset 6-7: nCols (2바이트)
        - offset 8-9: CellSpacing (2바이트)
        - ...

        Returns:
            (rows, cols) 튜플
        """
        if len(data) < 8:
            return (1, 1)

        rows = int.from_bytes(data[4:6], "little")
        cols = int.from_bytes(data[6:8], "little")

        # 최소값 보정
        rows = max(1, rows)
        cols = max(1, cols)

        return (rows, cols)

    def _decode_para_text(self, data: bytes) -> str:
        """PARA_TEXT 레코드 디코딩

        HWP 텍스트는 UTF-16LE로 인코딩되며,
        제어 문자(0x0001-0x001F)는 특수 의미를 가짐
        """
        if not data:
            return ""

        try:
            text = data.decode(DEFAULT_ENCODING, errors="replace")
        except Exception:
            return ""

        # 제어 문자 처리
        result = []
        i = 0
        text_len = len(text)

        # 모든 제어 문자 (탭, 줄바꿈, 단락구분 제외)
        FILTER_CHARS = set(range(0x0001, 0x0020)) - {0x000A, 0x000D, 0x0009}

        # 확장 제어 문자 (뒤에 8글자 추가 데이터 있음)
        EXTENDED_CTRL_CHARS = {0x0001, 0x0002, 0x0003, 0x000B, 0x000C,
                               0x0004, 0x0005, 0x0006, 0x0007, 0x0008,
                               0x0015, 0x0016, 0x0017, 0x0018, 0x0019}

        while i < text_len:
            char = text[i]
            code = ord(char)

            if code == CTRL_CHAR_PARA_BREAK:
                result.append("\n")
                i += 1
            elif code == CTRL_CHAR_LINE_BREAK:
                result.append("\n")
                i += 1
            elif code == 0x0009:
                result.append("\t")
                i += 1
            elif code in EXTENDED_CTRL_CHARS:
                # 확장 제어 문자: 현재 + 7글자(총 8글자/16바이트) 추가 데이터 스킵
                i += 8
            elif code in FILTER_CHARS:
                # 기타 제어 문자: 해당 문자만 스킵
                i += 1
            elif code >= 0x0020:
                result.append(char)
                i += 1
            else:
                i += 1

        return "".join(result).strip()


def extract_hwp5_structure(sections: list[tuple[str, bytes]]) -> DocumentStructure:
    """HWP 5.x 문서 구조 추출

    Args:
        sections: [(섹션이름, 압축해제된 데이터), ...]

    Returns:
        DocumentStructure 객체
    """
    doc = DocumentStructure()

    for idx, (section_name, section_data) in enumerate(sections):
        parser = StructureParser()
        section = parser.parse_section(idx, section_data)
        doc.sections.append(section)

    return doc
