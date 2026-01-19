"""HWP 파서 유틸리티"""

import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path


def convert_table_tags_to_markdown(text: str) -> str:
    """
    HWP 표 태그(<>)를 마크다운 테이블로 변환

    입력: <구 분><계약일자><계약자>
          <보험><2003.6.20><홍길동>

    출력: | 구 분 | 계약일자 | 계약자 |
          |-------|----------|--------|
          | 보험 | 2003.6.20 | 홍길동 |

    Args:
        text: 원본 텍스트

    Returns:
        변환된 텍스트
    """
    lines = text.split("\n")
    result = []
    table_rows = []
    in_table = False

    for line in lines:
        # 표 행 감지: <...><...> 패턴
        if re.search(r"<[^>]+>.*<[^>]+>", line):
            # 셀 추출
            cells = re.findall(r"<([^>]*)>", line)
            if cells:
                table_rows.append(cells)
                in_table = True
        else:
            # 표 종료 시 마크다운으로 변환
            if in_table and table_rows:
                result.extend(_rows_to_markdown(table_rows))
                table_rows = []
                in_table = False
            result.append(line)

    # 마지막 표 처리
    if table_rows:
        result.extend(_rows_to_markdown(table_rows))

    return "\n".join(result)


def _rows_to_markdown(rows: list[list[str]]) -> list[str]:
    """표 행을 마크다운으로 변환"""
    if not rows:
        return []

    # 열 너비 계산
    col_count = max(len(row) for row in rows)
    col_widths = [0] * col_count

    for row in rows:
        for i, cell in enumerate(row):
            if i < col_count:
                col_widths[i] = max(col_widths[i], len(cell.strip()))

    # 마크다운 생성
    md_lines = []
    for i, row in enumerate(rows):
        # 셀 패딩
        padded = []
        for j in range(col_count):
            cell = row[j].strip() if j < len(row) else ""
            padded.append(cell.ljust(col_widths[j]))

        md_lines.append("| " + " | ".join(padded) + " |")

        # 헤더 구분선 (첫 행 이후)
        if i == 0:
            separators = ["-" * max(3, w) for w in col_widths]
            md_lines.append("| " + " | ".join(separators) + " |")

    return md_lines


def extract_hwpx_text(filepath: str) -> str | None:
    """
    HWPX (XML 기반) 파일에서 텍스트 추출

    HWPX는 ZIP 압축된 XML 파일 모음

    Args:
        filepath: HWPX 파일 경로

    Returns:
        추출된 텍스트 또는 None
    """
    try:
        with zipfile.ZipFile(filepath, "r") as zf:
            namelist = zf.namelist()

            # 1. Preview/PrvText.txt 시도 (가장 쉬운 방법)
            prvtext_paths = [n for n in namelist if "prvtext" in n.lower()]
            for prvtext_path in prvtext_paths:
                try:
                    with zf.open(prvtext_path) as f:
                        text = f.read().decode("utf-8", errors="replace").strip()
                        if text:
                            return text
                except Exception:
                    continue

            # 2. Contents/section*.xml 파일에서 텍스트 추출
            texts = []
            section_files = sorted([
                n for n in namelist
                if "section" in n.lower() and n.endswith(".xml")
            ])

            for name in section_files:
                try:
                    with zf.open(name) as f:
                        content = f.read().decode("utf-8")
                        section_text = _parse_hwpx_section(content)
                        if section_text:
                            texts.append(section_text)
                except Exception:
                    continue

            return "\n\n".join(texts) if texts else None

    except (zipfile.BadZipFile, KeyError, Exception):
        return None


def _parse_hwpx_section(xml_content: str) -> str:
    """HWPX 섹션 XML 파싱"""
    try:
        # 네임스페이스 제거 (단순화)
        xml_content = re.sub(r'\sxmlns[^"]*"[^"]*"', "", xml_content)

        root = ET.fromstring(xml_content)
        texts = []

        # 모든 텍스트 요소 추출
        for elem in root.iter():
            if elem.text and elem.text.strip():
                texts.append(elem.text.strip())
            if elem.tail and elem.tail.strip():
                texts.append(elem.tail.strip())

        return "\n".join(texts)

    except ET.ParseError:
        return ""


def is_hwpx(filepath: str) -> bool:
    """
    파일이 HWPX 형식인지 확인

    HWPX는 ZIP 압축된 XML 파일 모음으로,
    mimetype 파일 또는 Contents/header.xml이 존재
    """
    try:
        with open(filepath, "rb") as f:
            # ZIP 시그니처 확인
            signature = f.read(4)
            if signature != b"PK\x03\x04":
                return False

        # HWPX 특정 파일 확인
        with zipfile.ZipFile(filepath, "r") as zf:
            namelist = zf.namelist()
            # HWPX는 mimetype 또는 Contents/header.xml 포함
            if "mimetype" in namelist:
                with zf.open("mimetype") as f:
                    mimetype = f.read().decode("utf-8", errors="ignore").strip()
                    if "hwp" in mimetype.lower():
                        return True
            if "Contents/header.xml" in namelist:
                return True
            # section0.xml도 HWPX 특징
            if any("section" in n and n.endswith(".xml") for n in namelist):
                return True

        return False

    except Exception:
        return False


def clean_text(text: str) -> str:
    """
    추출된 텍스트 정리

    - 연속 공백 제거
    - 연속 빈 줄 제거
    - 특수 제어 문자 제거
    """
    # 연속 공백 → 단일 공백
    text = re.sub(r"[ \t]+", " ", text)

    # 연속 빈 줄 → 단일 빈 줄
    text = re.sub(r"\n{3,}", "\n\n", text)

    # 특수 제어 문자 제거
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", text)

    return text.strip()
