"""HWP 파서 유틸리티"""

import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import logging

logger = logging.getLogger(__name__)


def _sort_section_files(files: list[str]) -> list[str]:
    """
    섹션 파일 숫자 기준 정렬 (Fix #4)

    문제: sorted()는 "section10.xml"을 "section2.xml" 앞에 배치
    해결: 파일명에서 숫자 추출하여 정수로 정렬

    Args:
        files: 섹션 파일 경로 리스트

    Returns:
        숫자 기준 정렬된 리스트
    """
    def extract_section_num(filepath: str) -> int:
        """파일명에서 섹션 번호 추출"""
        # "Contents/section0.xml" → 0, "section10.xml" → 10
        match = re.search(r'section(\d+)\.xml', filepath, re.IGNORECASE)
        if match:
            return int(match.group(1))
        return 0  # 숫자 없으면 맨 앞으로

    return sorted(files, key=extract_section_num)


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
            section_files = _sort_section_files([
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
    """HWPX 섹션 XML 파싱 (Fix #5: namespace-safe)"""
    try:
        # 방법 1: 네임스페이스 유지 파싱 시도
        root = ET.fromstring(xml_content)
        texts = _extract_text_preserving_whitespace(root)
        if texts:
            return "\n".join(texts)

    except ET.ParseError as e:
        logger.debug(f"Namespace-aware parsing failed: {e}, trying fallback")

    # 방법 2: 네임스페이스 제거 후 재시도 (fallback)
    try:
        xml_clean = _strip_namespaces_safe(xml_content)
        root = ET.fromstring(xml_clean)
        texts = _extract_text_preserving_whitespace(root)
        return "\n".join(texts)

    except ET.ParseError as e:
        logger.warning(f"HWPX section parsing failed: {e}")
        return ""


def _extract_text_preserving_whitespace(element: ET.Element) -> list[str]:
    """
    XML 요소에서 텍스트 추출 (whitespace 의미 보존)

    Fix #5: xml:space="preserve" 속성 존중
    """
    texts = []

    # xml:space 속성 확인
    space_preserve = element.get('{http://www.w3.org/XML/1998/namespace}space') == 'preserve'

    if element.text:
        text = element.text if space_preserve else element.text.strip()
        if text:
            texts.append(text)

    for child in element:
        child_texts = _extract_text_preserving_whitespace(child)
        texts.extend(child_texts)

    if element.tail:
        tail = element.tail if space_preserve else element.tail.strip()
        if tail:
            texts.append(tail)

    return texts


def _strip_namespaces_safe(xml_content: str) -> str:
    """
    네임스페이스 안전하게 제거 (Fix #5 fallback)

    주의: 이 방법은 의미를 잃을 수 있으므로 fallback으로만 사용
    """
    # xmlns 선언 제거
    xml_clean = re.sub(r'\sxmlns[^=]*="[^"]*"', "", xml_content)
    # 태그 접두사 제거 (예: <hp:p> → <p>)
    xml_clean = re.sub(r'<(/?)(\w+):', r'<\1', xml_clean)
    return xml_clean


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


# ============================================================================
# HWPX 구조 보존 추출 (XML 기반)
# ============================================================================

# HWPX XML 네임스페이스
HWPX_NAMESPACES = {
    "hp": "http://www.hancom.co.kr/hwpml/2011/paragraph",
    "hc": "http://www.hancom.co.kr/hwpml/2011/core",
    "hs": "http://www.hancom.co.kr/hwpml/2011/section",
}


def extract_hwpx_structure(filepath: str) -> dict | None:
    """
    HWPX 파일에서 구조 보존 추출

    문서 구조(단락, 테이블, 섹션)를 유지하여 딕셔너리로 반환

    Args:
        filepath: HWPX 파일 경로

    Returns:
        구조화된 딕셔너리 또는 None
        {
            "sections": [
                {
                    "index": 0,
                    "paragraphs": [{"text": "...", "level": 0}, ...],
                    "tables": [{"rows": 3, "cols": 2, "data": [[...], ...]}, ...]
                }
            ]
        }
    """
    if not is_hwpx(filepath):
        return None

    try:
        with zipfile.ZipFile(filepath, "r") as zf:
            namelist = zf.namelist()

            # section*.xml 파일 찾기 (숫자 기준 정렬)
            section_files = _sort_section_files([
                n for n in namelist
                if "section" in n.lower() and n.endswith(".xml")
            ])

            if not section_files:
                return None

            sections = []
            for idx, section_file in enumerate(section_files):
                try:
                    with zf.open(section_file) as f:
                        xml_content = f.read().decode("utf-8")
                        section_data = _parse_hwpx_section_structure(xml_content, idx)
                        if section_data:
                            sections.append(section_data)
                except Exception:
                    continue

            if not sections:
                return None

            return {"sections": sections}

    except (zipfile.BadZipFile, KeyError, Exception):
        return None


def _parse_hwpx_section_structure(xml_content: str, section_idx: int) -> dict | None:
    """
    HWPX 섹션 XML에서 구조 추출 (Fix #5: namespace-safe)

    Args:
        xml_content: section*.xml 내용
        section_idx: 섹션 인덱스

    Returns:
        섹션 구조 딕셔너리
    """
    paragraphs = []
    tables = []

    # 방법 1: 네임스페이스 유지 파싱 시도
    try:
        root = ET.fromstring(xml_content)
        _extract_structure_recursive(root, paragraphs, tables)

        if paragraphs or tables:
            return {
                "index": section_idx,
                "paragraphs": paragraphs,
                "tables": tables,
            }

    except ET.ParseError as e:
        logger.debug(f"Namespace-aware structure parsing failed: {e}, trying fallback")

    # 방법 2: 네임스페이스 제거 후 재시도 (fallback)
    try:
        xml_clean = _strip_namespaces_safe(xml_content)
        root = ET.fromstring(xml_clean)

        paragraphs = []
        tables = []
        _extract_structure_recursive(root, paragraphs, tables)

        return {
            "index": section_idx,
            "paragraphs": paragraphs,
            "tables": tables,
        }

    except ET.ParseError as e:
        logger.warning(f"HWPX section structure parsing failed: {e}")
        return None


def _get_local_tag(element: ET.Element) -> str:
    """
    네임스페이스 제거한 로컬 태그명 추출 (Fix #5)

    예: "{http://example.com}p" → "p"
    """
    tag = element.tag
    if tag.startswith("{"):
        # 네임스페이스 제거
        return tag.split("}", 1)[1].lower()
    return tag.lower()


def _extract_structure_recursive(
    element: ET.Element,
    paragraphs: list,
    tables: list,
    current_level: int = 0
) -> None:
    """
    XML 요소를 재귀적으로 탐색하여 단락과 테이블 추출 (Fix #5: namespace-safe)

    HWPX 태그 구조:
    - <p> or <para>: 단락
    - <t>: 텍스트
    - <tbl>: 테이블
    - <tr>: 테이블 행
    - <tc>: 테이블 셀
    - <run>: 텍스트 런
    """
    tag = _get_local_tag(element)

    # 테이블 처리
    if tag in ("tbl", "table"):
        table_data = _parse_hwpx_table(element)
        if table_data:
            tables.append(table_data)
        return  # 테이블 내부는 별도 처리

    # 단락 처리
    if tag in ("p", "para", "paragraph"):
        para_text = _extract_paragraph_text(element)
        if para_text.strip():
            paragraphs.append({
                "text": para_text.strip(),
                "level": current_level,
            })
        return  # 단락 내부 더 탐색 불필요

    # 섹션/컨테이너는 재귀 탐색
    for child in element:
        _extract_structure_recursive(child, paragraphs, tables, current_level)


def _extract_paragraph_text(element: ET.Element) -> str:
    """
    단락 요소에서 모든 텍스트 추출 (Fix #5: namespace-safe)

    <p>
      <run><t>텍스트1</t></run>
      <run><t>텍스트2</t></run>
    </p>
    """
    texts = []

    # 요소 자체의 텍스트
    if element.text and element.text.strip():
        texts.append(element.text.strip())

    # 모든 하위 요소 순회
    for elem in element.iter():
        tag = _get_local_tag(elem)

        # <t> 태그는 텍스트 컨테이너
        if tag == "t":
            if elem.text:
                texts.append(elem.text)

        # tail 텍스트 (닫는 태그 뒤 텍스트)
        if elem.tail and elem.tail.strip():
            texts.append(elem.tail.strip())

    return "".join(texts)


def _parse_hwpx_table(table_element: ET.Element) -> dict | None:
    """
    테이블 요소에서 구조 추출 (Fix #5: namespace-safe)

    <tbl>
      <tr><tc><p>...</p></tc><tc><p>...</p></tc></tr>
      <tr><tc><p>...</p></tc><tc><p>...</p></tc></tr>
    </tbl>
    """
    rows_data = []

    # 행 찾기 (<tr>)
    for child in table_element:
        tag = _get_local_tag(child)
        if tag == "tr":
            row_cells = _parse_hwpx_table_row(child)
            if row_cells:
                rows_data.append(row_cells)

    if not rows_data:
        return None

    # 열 수 정규화 (가장 긴 행 기준)
    max_cols = max(len(row) for row in rows_data)
    for row in rows_data:
        while len(row) < max_cols:
            row.append("")

    return {
        "rows": len(rows_data),
        "cols": max_cols,
        "data": rows_data,
    }


def _parse_hwpx_table_row(tr_element: ET.Element) -> list[str]:
    """
    테이블 행에서 셀 텍스트 추출 (Fix #5: namespace-safe)

    <tr>
      <tc><p><run><t>셀1</t></run></p></tc>
      <tc><p><run><t>셀2</t></run></p></tc>
    </tr>
    """
    cells = []

    for child in tr_element:
        tag = _get_local_tag(child)
        if tag == "tc":
            cell_text = _extract_cell_text(child)
            cells.append(cell_text.strip())

    return cells


def _extract_cell_text(tc_element: ET.Element) -> str:
    """
    테이블 셀에서 텍스트 추출 (Fix #5: namespace-safe)

    셀 안에는 여러 단락이 있을 수 있음
    """
    texts = []

    for elem in tc_element.iter():
        tag = _get_local_tag(elem)

        if tag == "t":
            if elem.text:
                texts.append(elem.text)

    return " ".join(texts)


def hwpx_structure_to_flat_text(structure: dict) -> str:
    """
    HWPX 구조를 평탄화된 텍스트로 변환 (하위 호환)

    Args:
        structure: extract_hwpx_structure() 결과

    Returns:
        평탄화된 텍스트
    """
    texts = []

    for section in structure.get("sections", []):
        for para in section.get("paragraphs", []):
            text = para.get("text", "").strip()
            if text:
                texts.append(text)

    return "\n".join(texts)
