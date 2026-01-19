"""HWP 텍스트 추출기"""

import os
from typing import Literal

from .constants import STREAM_PRV_TEXT, DEFAULT_ENCODING
from .reader import HWPReader, HWPReaderError
from .record import RecordParser
from .models import ExtractResult


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


def extract_hwp_text(filepath: str) -> ExtractResult:
    """
    HWP 파일에서 텍스트 추출 (편의 함수)

    Args:
        filepath: HWP 파일 경로

    Returns:
        ExtractResult 객체
    """
    try:
        with HWPReader(filepath) as reader:
            extractor = TextExtractor(reader)
            return extractor.extract()
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
