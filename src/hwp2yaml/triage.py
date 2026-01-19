"""HWP 파일 트리아지 (버전 감지 및 분류)"""

import os
from enum import Enum
from dataclasses import dataclass
from typing import Callable

from .utils import is_hwpx


class HWPVersion(Enum):
    """HWP 파일 버전"""
    HWP_3X = "hwp3"      # HWP 3.x (1990년대, 구 형식)
    HWP_5X = "hwp5"      # HWP 5.x (OLE2 기반, 2000년대~)
    HWPX = "hwpx"        # HWPX (XML + ZIP, 2010년대~)
    UNKNOWN = "unknown"  # 알 수 없는 형식


# 파일 시그니처
OLE2_SIGNATURE = b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1"
ZIP_SIGNATURE = b"PK\x03\x04"
HWP3_SIGNATURE = b"HWP Document File"  # HWP 3.x 시작 부분


@dataclass
class TriageResult:
    """트리아지 결과"""
    filepath: str
    version: HWPVersion
    file_size: int
    can_process: bool
    note: str = ""


def detect_hwp_version(filepath: str) -> HWPVersion:
    """
    HWP 파일 버전 감지

    Args:
        filepath: HWP 파일 경로

    Returns:
        HWPVersion enum
    """
    if not os.path.isfile(filepath):
        return HWPVersion.UNKNOWN

    try:
        with open(filepath, "rb") as f:
            header = f.read(32)

        # 1. HWPX 확인 (ZIP 시그니처)
        if header[:4] == ZIP_SIGNATURE:
            if is_hwpx(filepath):
                return HWPVersion.HWPX
            return HWPVersion.UNKNOWN

        # 2. OLE2 확인 (HWP 5.x)
        if header[:8] == OLE2_SIGNATURE:
            # OLE2 파일 내부에서 HWP 시그니처 확인
            try:
                import olefile
                ole = olefile.OleFileIO(filepath)
                if ole.exists("FileHeader"):
                    file_header = ole.openstream("FileHeader").read(32)
                    if file_header[:17] == HWP3_SIGNATURE:
                        ole.close()
                        return HWPVersion.HWP_5X
                ole.close()
            except Exception:
                pass
            return HWPVersion.HWP_5X  # OLE2이면 일단 5.x로 간주

        # 3. HWP 3.x 확인
        # HWP 3.x는 고유한 바이너리 포맷
        # 첫 바이트가 특정 패턴인지 확인
        if header[:17] == HWP3_SIGNATURE:
            return HWPVersion.HWP_3X

        # 파일 명령으로 추가 확인
        return _detect_by_file_command(filepath)

    except Exception:
        return HWPVersion.UNKNOWN


def _detect_by_file_command(filepath: str) -> HWPVersion:
    """file 명령으로 HWP 버전 감지 (폴백)"""
    import subprocess

    try:
        result = subprocess.run(
            ["file", "-b", filepath],
            capture_output=True,
            text=True,
            timeout=5,
        )
        output = result.stdout.lower()

        if "version 3" in output:
            return HWPVersion.HWP_3X
        elif "version 5" in output:
            return HWPVersion.HWP_5X
        elif "hwpx" in output:
            return HWPVersion.HWPX
        elif "hwp" in output:
            # 버전 불명확하면 OLE2 여부로 판단
            with open(filepath, "rb") as f:
                if f.read(8) == OLE2_SIGNATURE:
                    return HWPVersion.HWP_5X
            return HWPVersion.HWP_3X

    except Exception:
        pass

    return HWPVersion.UNKNOWN


def triage_file(filepath: str) -> TriageResult:
    """
    단일 파일 트리아지

    Args:
        filepath: HWP 파일 경로

    Returns:
        TriageResult
    """
    version = detect_hwp_version(filepath)
    file_size = os.path.getsize(filepath) if os.path.isfile(filepath) else 0

    can_process = version in (HWPVersion.HWP_5X, HWPVersion.HWPX)

    notes = {
        HWPVersion.HWP_3X: "HWP 3.x 미지원 (구 형식)",
        HWPVersion.HWP_5X: "HWP 5.x 처리 가능",
        HWPVersion.HWPX: "HWPX 처리 가능",
        HWPVersion.UNKNOWN: "알 수 없는 형식",
    }

    return TriageResult(
        filepath=filepath,
        version=version,
        file_size=file_size,
        can_process=can_process,
        note=notes.get(version, ""),
    )


@dataclass
class TriageSummary:
    """트리아지 요약"""
    total: int
    hwp3_count: int
    hwp5_count: int
    hwpx_count: int
    unknown_count: int
    processable: int
    skipped: int

    hwp3_files: list[str]
    hwp5_files: list[str]
    hwpx_files: list[str]
    unknown_files: list[str]


def triage_files(filepaths: list[str], progress: bool = True) -> TriageSummary:
    """
    파일 목록 트리아지

    Args:
        filepaths: HWP 파일 경로 목록
        progress: 진행률 표시

    Returns:
        TriageSummary
    """
    hwp3_files = []
    hwp5_files = []
    hwpx_files = []
    unknown_files = []

    iterator = filepaths
    if progress:
        try:
            from tqdm import tqdm
            iterator = tqdm(filepaths, desc="트리아지")
        except ImportError:
            pass

    for filepath in iterator:
        result = triage_file(filepath)

        if result.version == HWPVersion.HWP_3X:
            hwp3_files.append(filepath)
        elif result.version == HWPVersion.HWP_5X:
            hwp5_files.append(filepath)
        elif result.version == HWPVersion.HWPX:
            hwpx_files.append(filepath)
        else:
            unknown_files.append(filepath)

    return TriageSummary(
        total=len(filepaths),
        hwp3_count=len(hwp3_files),
        hwp5_count=len(hwp5_files),
        hwpx_count=len(hwpx_files),
        unknown_count=len(unknown_files),
        processable=len(hwp5_files) + len(hwpx_files),
        skipped=len(hwp3_files) + len(unknown_files),
        hwp3_files=hwp3_files,
        hwp5_files=hwp5_files,
        hwpx_files=hwpx_files,
        unknown_files=unknown_files,
    )


def triage_directory(directory: str, recursive: bool = True, progress: bool = True) -> TriageSummary:
    """
    디렉토리 트리아지

    Args:
        directory: 디렉토리 경로
        recursive: 하위 디렉토리 포함
        progress: 진행률 표시

    Returns:
        TriageSummary
    """
    import glob

    pattern = "**/*.hwp" if recursive else "*.hwp"
    filepaths = glob.glob(os.path.join(directory, pattern), recursive=recursive)

    return triage_files(filepaths, progress=progress)
