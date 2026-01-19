"""HWP 파일 읽기 및 OLE 스트림 접근"""

import os
import zlib
from pathlib import Path
from typing import Iterator

import olefile

from .constants import (
    HWP_SIGNATURE,
    STREAM_FILE_HEADER,
    STREAM_PRV_TEXT,
    STREAM_BODY_TEXT,
    FLAG_COMPRESSED,
    FLAG_ENCRYPTED,
    MAX_FILE_SIZE_MB,
)
from .models import HWPMetadata, HWPVersion


class HWPReaderError(Exception):
    """HWP 읽기 오류"""
    pass


class HWPReader:
    """HWP 파일 읽기 클래스"""

    def __init__(self, filepath: str):
        """
        HWP 파일 열기

        Args:
            filepath: HWP 파일 경로

        Raises:
            HWPReaderError: 파일 열기 실패
        """
        self.filepath = os.path.realpath(filepath)  # 경로 정규화 (보안)
        self._ole: olefile.OleFileIO | None = None
        self._metadata: HWPMetadata | None = None
        self._file_header: bytes | None = None

        self._open()

    def _open(self) -> None:
        """OLE 파일 열기"""
        # 파일 존재 확인
        if not os.path.isfile(self.filepath):
            raise HWPReaderError(f"파일이 존재하지 않음: {self.filepath}")

        # 파일 크기 확인
        file_size = os.path.getsize(self.filepath)
        if file_size > MAX_FILE_SIZE_MB * 1024 * 1024:
            raise HWPReaderError(f"파일 크기 초과 ({file_size / 1024 / 1024:.1f}MB > {MAX_FILE_SIZE_MB}MB)")

        # OLE 파일 열기
        try:
            self._ole = olefile.OleFileIO(self.filepath)
        except Exception as e:
            raise HWPReaderError(f"OLE 파일 열기 실패: {e}")

        # HWP 유효성 확인
        if not self.is_valid_hwp():
            self.close()
            raise HWPReaderError("유효한 HWP 5.x 파일이 아님")

        # 암호화 확인
        if self.is_encrypted():
            self.close()
            raise HWPReaderError("암호화된 HWP 파일은 지원하지 않음")

        # 메타데이터 생성
        self._metadata = HWPMetadata(
            filepath=self.filepath,
            filename=os.path.basename(self.filepath),
            version=self.get_version(),
            is_compressed=self.is_compressed(),
            is_encrypted=False,
            file_size_bytes=file_size,
            streams=self.list_streams(),
        )

    def close(self) -> None:
        """파일 닫기"""
        if self._ole:
            self._ole.close()
            self._ole = None

    def __enter__(self) -> "HWPReader":
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        self.close()

    def _get_file_header(self) -> bytes:
        """FileHeader 스트림 읽기 (캐시)"""
        if self._file_header is None:
            try:
                self._file_header = self.open_stream(STREAM_FILE_HEADER)
            except Exception:
                self._file_header = b""
        return self._file_header

    def is_valid_hwp(self) -> bool:
        """HWP 5.x 파일인지 확인"""
        if not self._ole:
            return False

        # FileHeader 스트림 존재 확인
        if not self._ole.exists(STREAM_FILE_HEADER):
            return False

        # 시그니처 확인
        header = self._get_file_header()
        return header[:len(HWP_SIGNATURE)] == HWP_SIGNATURE

    def is_encrypted(self) -> bool:
        """암호화 여부 확인"""
        header = self._get_file_header()
        if len(header) < 40:
            return False

        # offset 36: 속성 플래그 (4바이트, little-endian)
        flags = int.from_bytes(header[36:40], "little")
        return bool(flags & FLAG_ENCRYPTED)

    def is_compressed(self) -> bool:
        """압축 여부 확인"""
        header = self._get_file_header()
        if len(header) < 40:
            return True  # 기본값: 압축

        flags = int.from_bytes(header[36:40], "little")
        return bool(flags & FLAG_COMPRESSED)

    def get_version(self) -> HWPVersion | None:
        """HWP 버전 정보"""
        header = self._get_file_header()
        if len(header) < 36:
            return None

        # offset 32: 버전 (4바이트)
        # major.minor.build.revision (각 1바이트)
        version_bytes = header[32:36]
        return HWPVersion(
            major=version_bytes[3],
            minor=version_bytes[2],
            build=version_bytes[1],
            revision=version_bytes[0],
        )

    def list_streams(self) -> list[str]:
        """OLE 스트림 목록"""
        if not self._ole:
            return []

        streams = []
        for entry in self._ole.listdir():
            streams.append("/".join(entry))
        return streams

    def has_stream(self, name: str) -> bool:
        """스트림 존재 여부"""
        if not self._ole:
            return False
        return self._ole.exists(name)

    def open_stream(self, name: str) -> bytes:
        """
        스트림 데이터 읽기

        Args:
            name: 스트림 이름 (예: "PrvText", "BodyText/Section0")

        Returns:
            스트림 데이터 (바이트)

        Raises:
            HWPReaderError: 스트림 읽기 실패
        """
        if not self._ole:
            raise HWPReaderError("파일이 열려있지 않음")

        if not self._ole.exists(name):
            raise HWPReaderError(f"스트림이 존재하지 않음: {name}")

        try:
            return self._ole.openstream(name).read()
        except Exception as e:
            raise HWPReaderError(f"스트림 읽기 실패 ({name}): {e}")

    def open_stream_decompressed(self, name: str) -> bytes:
        """
        압축된 스트림 읽기 및 해제

        Args:
            name: 스트림 이름

        Returns:
            압축 해제된 데이터
        """
        data = self.open_stream(name)

        if not self.is_compressed():
            return data

        # zlib 압축 해제 (-15: raw deflate, no header)
        try:
            return zlib.decompress(data, -15)
        except zlib.error:
            # 압축되지 않은 데이터일 수 있음
            return data

    def iter_sections(self) -> Iterator[tuple[str, bytes]]:
        """
        BodyText 섹션 순회

        Yields:
            (섹션 이름, 압축 해제된 데이터)
        """
        if not self._ole:
            return

        section_idx = 0
        while True:
            section_name = f"{STREAM_BODY_TEXT}/Section{section_idx}"
            if not self.has_stream(section_name):
                break

            try:
                data = self.open_stream_decompressed(section_name)
                yield section_name, data
            except HWPReaderError:
                break

            section_idx += 1

    @property
    def metadata(self) -> HWPMetadata | None:
        """문서 메타데이터"""
        return self._metadata
