"""HWP 텍스트 추출기"""

import os
from typing import Literal

from .constants import STREAM_PRV_TEXT, DEFAULT_ENCODING
from .reader import HWPReader, HWPReaderError
from .record import RecordParser
from .models import ExtractResult, StructuredResult
from .utils import (
    is_hwpx,
    extract_hwpx_text,
    extract_hwpx_structure,
    hwpx_structure_to_flat_text,
    convert_table_tags_to_markdown,
    clean_text,
)


class TextExtractor:
    """
    HWP 텍스트 추출기

    추출 전략:
    1. PrvText 스트림 시도 (빠르고 간단)
    2. PrvText 없으면 BodyText 파싱 (완전하지만 복잡)
    """

    def __init__(self, reader: HWPReader):
        """
        Args:
            reader: HWPReader 인스턴스
        """
        self.reader = reader

    def extract(self) -> ExtractResult:
        """
        텍스트 추출 실행

        Returns:
            ExtractResult 객체
        """
        filepath = self.reader.filepath

        # 1. PrvText 시도
        text = self._extract_prvtext()
        if text:
            return ExtractResult(
                filepath=filepath,
                success=True,
                text=text,
                method="prvtext",
                metadata=self.reader.metadata,
            )

        # 2. BodyText 폴백
        text = self._extract_bodytext()
        if text:
            return ExtractResult(
                filepath=filepath,
                success=True,
                text=text,
                method="bodytext",
                metadata=self.reader.metadata,
            )

        # 3. 추출 실패
        return ExtractResult(
            filepath=filepath,
            success=False,
            text=None,
            method="failed",
            error="텍스트 추출 실패: PrvText 없음, BodyText 파싱 실패",
            metadata=self.reader.metadata,
        )

    def _extract_prvtext(self) -> str | None:
        """
        PrvText 스트림에서 미리보기 텍스트 추출

        PrvText는 압축되지 않은 UTF-16LE 텍스트로,
        문서의 첫 부분(~4KB)을 포함

        Returns:
            추출된 텍스트 또는 None
        """
        if not self.reader.has_stream(STREAM_PRV_TEXT):
            return None

        try:
            data = self.reader.open_stream(STREAM_PRV_TEXT)
        except HWPReaderError:
            return None

        if not data:
            return None

        # UTF-16LE 디코딩
        try:
            text = data.decode(DEFAULT_ENCODING, errors="replace")
        except Exception:
            return None

        # 널 문자 및 공백 정리
        text = text.replace("\x00", "").strip()

        if not text:
            return None

        return text

    def _extract_bodytext(self) -> str | None:
        """
        BodyText 섹션에서 전체 본문 추출

        Returns:
            추출된 텍스트 또는 None
        """
        sections = list(self.reader.iter_sections())

        if not sections:
            return None

        try:
            text = RecordParser.extract_all_text(sections)
        except Exception:
            return None

        if not text:
            return None

        return text


def extract_hwp_text(
    filepath: str,
    convert_tables: bool = True,
    clean: bool = True,
) -> ExtractResult:
    """
    HWP/HWPX 파일에서 텍스트 추출 (편의 함수)

    Args:
        filepath: HWP/HWPX 파일 경로
        convert_tables: 표 태그를 마크다운으로 변환
        clean: 텍스트 정리 (공백, 제어문자)

    Returns:
        ExtractResult 객체
    """
    # 1. HWPX (XML 기반) 먼저 확인
    if is_hwpx(filepath):
        text = extract_hwpx_text(filepath)
        if text:
            if convert_tables:
                text = convert_table_tags_to_markdown(text)
            if clean:
                text = clean_text(text)

            return ExtractResult(
                filepath=filepath,
                success=True,
                text=text,
                method="hwpx",
            )
        else:
            return ExtractResult(
                filepath=filepath,
                success=False,
                text=None,
                method="failed",
                error="HWPX 파싱 실패",
            )

    # 2. HWP (OLE2 기반)
    try:
        with HWPReader(filepath) as reader:
            extractor = TextExtractor(reader)
            result = extractor.extract()

            # 후처리
            if result.success and result.text:
                text = result.text
                if convert_tables:
                    text = convert_table_tags_to_markdown(text)
                if clean:
                    text = clean_text(text)
                result.text = text
                result.char_count = len(text)

            return result

    except HWPReaderError as e:
        return ExtractResult(
            filepath=filepath,
            success=False,
            text=None,
            method="failed",
            error=str(e),
        )
    except Exception as e:
        return ExtractResult(
            filepath=filepath,
            success=False,
            text=None,
            method="failed",
            error=f"예상치 못한 오류: {e}",
        )


def extract_hwp_structure(filepath: str) -> StructuredResult:
    """
    HWP/HWPX 파일에서 구조 보존 추출

    단락, 테이블, 섹션 구조를 유지하여 추출

    Args:
        filepath: HWP/HWPX 파일 경로

    Returns:
        StructuredResult 객체
    """
    # HWPX (XML 기반) 처리
    if is_hwpx(filepath):
        structure = extract_hwpx_structure(filepath)

        if structure:
            # 테이블 추출
            tables = []
            for section in structure.get("sections", []):
                tables.extend(section.get("tables", []))

            # 평탄화된 텍스트 (하위 호환)
            flat_text = hwpx_structure_to_flat_text(structure)

            return StructuredResult(
                filepath=filepath,
                success=True,
                method="hwpx_structure",
                structure=structure,
                text=flat_text,
                tables=tables,
            )
        else:
            return StructuredResult(
                filepath=filepath,
                success=False,
                method="failed",
                error="HWPX 구조 파싱 실패",
            )

    # HWP 5.x (OLE2 기반) 처리
    from .structure import extract_hwp5_structure

    try:
        with HWPReader(filepath) as reader:
            sections = list(reader.iter_sections())

            if not sections:
                return StructuredResult(
                    filepath=filepath,
                    success=False,
                    method="failed",
                    error="BodyText 섹션 없음",
                    metadata=reader.metadata,
                )

            # 구조 파싱
            doc_structure = extract_hwp5_structure(sections)

            # 테이블 추출
            tables = [t.to_dict() for t in doc_structure.get_all_tables()]

            # 평탄화된 텍스트 (하위 호환)
            flat_text = doc_structure.to_flat_text()

            return StructuredResult(
                filepath=filepath,
                success=True,
                method="hwp5_structure",
                structure=doc_structure.to_dict(),
                text=flat_text,
                tables=tables,
                metadata=reader.metadata,
            )

    except HWPReaderError as e:
        return StructuredResult(
            filepath=filepath,
            success=False,
            method="failed",
            error=str(e),
        )
    except Exception as e:
        return StructuredResult(
            filepath=filepath,
            success=False,
            method="failed",
            error=f"구조 추출 오류: {e}",
        )
