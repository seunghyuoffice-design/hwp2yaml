"""HWP 레코드 구조 파싱"""

from typing import Iterator
from dataclasses import dataclass

from .constants import (
    HWPTAG_PARA_TEXT,
    HWPTAG_PARA_HEADER,
    CTRL_CHAR_PARA_BREAK,
    CTRL_CHAR_LINE_BREAK,
    DEFAULT_ENCODING,
)
from .models import RecordHeader


@dataclass
class Record:
    """HWP 레코드"""
    header: RecordHeader
    data: bytes


class RecordParser:
    """HWP 레코드 파서"""

    # 텍스트에서 필터링할 제어 문자 범위
    FILTER_CHARS = set(range(0x0001, 0x0020)) - {
        CTRL_CHAR_LINE_BREAK,   # 0x000A: 줄바꿈 유지
        CTRL_CHAR_PARA_BREAK,   # 0x000D: 단락 구분 유지
        0x0009,                 # 탭 유지
    }

    @staticmethod
    def parse_header(data: bytes, offset: int = 0) -> tuple[RecordHeader, int]:
        """
        레코드 헤더 파싱

        HWP 레코드 헤더 구조 (4바이트):
        - bits 0-9: TagID (10비트)
        - bits 10-19: Level (10비트)
        - bits 20-31: Size (12비트)

        Size가 0xFFF(4095)이면 추가 4바이트에 실제 크기 저장

        Args:
            data: 바이트 데이터
            offset: 시작 오프셋

        Returns:
            (RecordHeader, 다음 오프셋)
        """
        if len(data) < offset + 4:
            raise ValueError("데이터 부족: 레코드 헤더 파싱 불가")

        # 4바이트 리틀 엔디안
        header_val = int.from_bytes(data[offset:offset + 4], "little")

        tag_id = header_val & 0x3FF           # 하위 10비트
        level = (header_val >> 10) & 0x3FF    # 다음 10비트
        size = (header_val >> 20) & 0xFFF     # 상위 12비트

        next_offset = offset + 4

        # 확장 크기 (Size == 0xFFF)
        if size == 0xFFF:
            if len(data) < next_offset + 4:
                raise ValueError("데이터 부족: 확장 크기 파싱 불가")
            size = int.from_bytes(data[next_offset:next_offset + 4], "little")
            next_offset += 4

        return RecordHeader(tag_id=tag_id, level=level, size=size), next_offset

    @classmethod
    def iter_records(cls, data: bytes) -> Iterator[Record]:
        """
        섹션 데이터에서 레코드 순회

        Args:
            data: 압축 해제된 섹션 데이터

        Yields:
            Record 객체
        """
        offset = 0
        data_len = len(data)

        while offset < data_len:
            try:
                header, next_offset = cls.parse_header(data, offset)
            except ValueError:
                break

            # 레코드 데이터 추출
            record_end = next_offset + header.size
            if record_end > data_len:
                # 데이터 부족: 남은 부분만 사용
                record_data = data[next_offset:]
            else:
                record_data = data[next_offset:record_end]

            yield Record(header=header, data=record_data)

            offset = next_offset + header.size

    @classmethod
    def extract_text_from_section(cls, section_data: bytes) -> str:
        """
        섹션 데이터에서 텍스트 추출

        Args:
            section_data: 압축 해제된 섹션 데이터

        Returns:
            추출된 텍스트
        """
        texts = []

        for record in cls.iter_records(section_data):
            if record.header.tag_id == HWPTAG_PARA_TEXT:
                text = cls._decode_para_text(record.data)
                if text:
                    texts.append(text)

        return "\n".join(texts)

    @classmethod
    def _decode_para_text(cls, data: bytes) -> str:
        """
        PARA_TEXT 레코드 디코딩

        HWP 텍스트는 UTF-16LE로 인코딩되며,
        제어 문자(0x0001-0x001F)는 특수 의미를 가짐

        Args:
            data: PARA_TEXT 레코드 데이터

        Returns:
            디코딩된 텍스트
        """
        if not data:
            return ""

        # UTF-16LE 디코딩
        try:
            text = data.decode(DEFAULT_ENCODING, errors="replace")
        except Exception:
            return ""

        # 제어 문자 처리
        result = []
        i = 0
        text_len = len(text)

        while i < text_len:
            char = text[i]
            code = ord(char)

            if code == CTRL_CHAR_PARA_BREAK:
                # 단락 구분 → 줄바꿈
                result.append("\n")
            elif code == CTRL_CHAR_LINE_BREAK:
                # 줄바꿈 (표 내)
                result.append("\n")
            elif code == 0x0009:
                # 탭
                result.append("\t")
            elif code in cls.FILTER_CHARS:
                # 확장 제어 문자: 스킵
                # 일부는 추가 데이터를 가짐 (inline 객체)
                if code in (0x0001, 0x0002, 0x0003, 0x000B, 0x000C):
                    # 8글자 추가 데이터 스킵
                    i += 8
            elif code >= 0x0020:
                # 일반 문자
                result.append(char)

            i += 1

        return "".join(result).strip()

    @classmethod
    def extract_all_text(cls, sections: list[tuple[str, bytes]]) -> str:
        """
        모든 섹션에서 텍스트 추출

        Args:
            sections: [(섹션이름, 데이터), ...]

        Returns:
            전체 텍스트
        """
        all_texts = []

        for section_name, section_data in sections:
            text = cls.extract_text_from_section(section_data)
            if text:
                all_texts.append(text)

        return "\n\n".join(all_texts)
