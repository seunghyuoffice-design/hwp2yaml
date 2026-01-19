# hwp2yaml

HWP 5.x / HWPX / HWP 3.x 문서를 YAML로 직접 변환하는 Python 라이브러리.

[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](https://opensource.org/licenses/MIT)

## 요구사항

- **Python**: 3.10 이상
- **OS**: Windows, macOS, Linux

## 빠른 시작

```bash
pip install hwp2yaml
```

```python
from hwp2yaml import extract_hwp_text

result = extract_hwp_text("document.hwp")
if result.success:
    print(result.text)
```

## 지원 형식

| 형식 | 지원 | 방식 | 비고 |
|------|------|------|------|
| HWP 5.x | ✅ | 직접 파싱 | OLE2 기반 (2002~) |
| HWPX | ✅ | 직접 파싱 | XML + ZIP 기반 (2014~) |
| HWP 3.x | ✅ | LibreOffice + Docling | 구형식 (1990년대) |

## 특징

- **HWP → YAML**: 중간 변환 없이 직접 YAML 출력
- **순수 Python**: pyhwp 등 AGPL 라이브러리 대체
- **MIT 라이선스**: 상업적 사용 가능
- **전 버전 지원**: HWP 3.x ~ HWPX 모두 처리
- **트리아지**: 파일 버전 자동 감지 및 분류
- **구조 보존**: 단락, 테이블, 섹션 구조 완전 보존 (HWP 5.x)

## 설치

### PyPI (권장)

```bash
pip install hwp2yaml
```

### 소스에서 설치

```bash
git clone https://github.com/seunghyuoffice-design/hwp2yaml.git
cd hwp2yaml
pip install -e .
```

### 개발 환경 설치

```bash
pip install -e ".[dev]"
```

### HWP 3.x 변환 의존성 (선택)

HWP 3.x 파일을 변환하려면 추가 설치가 필요합니다:

**Ubuntu/Debian:**
```bash
sudo apt install libreoffice poppler-utils
pip install docling
```

**macOS:**
```bash
brew install --cask libreoffice
brew install poppler
pip install docling
```

**Windows:**
1. [LibreOffice](https://www.libreoffice.org/download/download/) 설치
2. [Poppler for Windows](https://github.com/oschwartz10612/poppler-windows/releases) 설치 후 PATH에 추가
3. `pip install docling`

## CLI 사용법

```bash
# 단일 파일 변환
hwp2yaml document.hwp

# 출력 파일 지정
hwp2yaml document.hwp -o output.yaml

# 디렉토리 일괄 변환
hwp2yaml /path/to/files/ -o /path/to/output/
```

## API 사용법

### 1. 텍스트 추출 (가장 간단)

```python
from hwp2yaml import extract_hwp_text

result = extract_hwp_text("document.hwp")

if result.success:
    print(result.text)       # 추출된 텍스트
    print(result.method)     # "prvtext", "bodytext", "hwpx" 중 하나
else:
    print(f"실패: {result.error}")
```

### 2. 구조 보존 추출 (HWP 5.x)

단락, 테이블, 섹션 구조를 보존하여 YAML로 변환:

```python
from hwp2yaml import extract_hwp_structure

result = extract_hwp_structure("document.hwp")

if result.success:
    # 구조 정보 접근
    for section in result.structure["sections"]:
        print(f"단락 수: {len(section['paragraphs'])}")
        print(f"테이블 수: {len(section['tables'])}")

    # 테이블 데이터 접근
    for table in result.tables:
        print(f"테이블: {table['rows']}x{table['cols']}")
        for row in table['data']:
            print(row)

    # YAML 문자열로 출력
    print(result.to_yaml())

    # YAML 파일로 저장
    with open("output.yaml", "w", encoding="utf-8") as f:
        f.write(result.to_yaml())
```

### 3. 파일 버전 감지 (트리아지)

```python
from hwp2yaml import detect_hwp_version, HWPVersion, triage_directory

# 단일 파일 버전 확인
version = detect_hwp_version("document.hwp")

if version == HWPVersion.HWP_5X:
    print("HWP 5.x - 직접 처리 가능")
elif version == HWPVersion.HWPX:
    print("HWPX - 직접 처리 가능")
elif version == HWPVersion.HWP_3X:
    print("HWP 3.x - LibreOffice 필요")
elif version == HWPVersion.UNKNOWN:
    print("알 수 없는 형식")

# 디렉토리 일괄 트리아지
summary = triage_directory("/path/to/files")
print(f"총 파일: {summary.total}")
print(f"HWP 5.x: {summary.hwp5_count}")
print(f"HWPX: {summary.hwpx_count}")
print(f"HWP 3.x: {summary.hwp3_count}")
print(f"알 수 없음: {summary.unknown_count}")
```

### 4. HWP 3.x 변환

```python
from hwp2yaml import convert_hwp3

result = convert_hwp3("old_document.hwp")

if result.success:
    print(result.text)                    # 추출된 텍스트
    print(f"테이블 수: {len(result.tables)}")
    print(f"방법: {result.method}")       # "hwp3_docling" 또는 "hwp3_pdftotext"

    # YAML로 변환
    yaml_str = result.to_yaml()
    yaml_dict = result.to_yaml_dict()
```

### 5. 배치 처리

```python
from hwp2yaml import BatchProcessor, batch_convert_to_yaml

# 방법 1: BatchProcessor 클래스
processor = BatchProcessor(workers=4)
result = processor.process_directory("/path/to/files")
print(f"성공: {result.success}/{result.total}")

# 방법 2: 함수 직접 호출
yaml_dicts = batch_convert_to_yaml(
    filepaths=["file1.hwp", "file2.hwp", "file3.hwp"],
    output_dir="/path/to/yaml",           # 개별 YAML 파일 저장
    combined_output="/path/to/all.yaml",  # 통합 파일 저장 (선택)
)
```

## 반환 타입

### ExtractResult

텍스트 추출 결과를 담는 dataclass:

```python
@dataclass
class ExtractResult:
    success: bool          # 성공 여부
    text: str              # 추출된 텍스트 (실패 시 빈 문자열)
    method: str            # 사용된 방법 ("prvtext", "bodytext", "hwpx", "hwp3_docling", ...)
    error: str | None      # 에러 메시지 (성공 시 None)
    tables: list[dict]     # 추출된 테이블 목록
    structure: dict | None # 구조 정보 (extract_hwp_structure 사용 시)

    def to_yaml(self) -> str:
        """YAML 문자열로 변환"""

    def to_yaml_dict(self) -> dict:
        """YAML 호환 딕셔너리로 변환"""
```

### HWPVersion

파일 버전을 나타내는 Enum:

```python
class HWPVersion(Enum):
    HWP_3X = "hwp3"      # HWP 3.x (1990년대)
    HWP_5X = "hwp5"      # HWP 5.x (2002~)
    HWPX = "hwpx"        # HWPX (2014~)
    UNKNOWN = "unknown"  # 알 수 없음
```

## YAML 출력 스키마

### 구조 보존 스키마 (HWP 5.x)

```yaml
metadata:
  source_file: /path/to/document.hwp
  method: hwp5_structure
  extracted_at: "2026-01-19T12:34:56"

structure:
  sections:
    - index: 0
      paragraphs:
        - text: "첫 번째 단락"
          level: 0
        - text: "두 번째 단락"
          level: 0
      tables:
        - rows: 3
          cols: 2
          data:
            - ["헤더1", "헤더2"]
            - ["셀1", "셀2"]
            - ["셀3", "셀4"]

tables:  # 전체 테이블 목록 (편의용)
  - rows: 3
    cols: 2
    data: [...]

raw_text: |
  평탄화된 전체 텍스트
```

### HWP 3.x 변환 스키마

```yaml
metadata:
  source_file: /path/to/file.hwp
  version: hwp3
  method: hwp3_docling
  converted_at: "2026-01-19T12:34:56"

content:
  parties: 당사자 정보
  request: 신청취지
  facts: 사실관계
  decision: 위원회 판단

tables:
  - table_id: table_0
    rows: 3
    cols: 2
    cells:
      - {row: 0, col: 0, text: "..."}

raw_text: |
  원본 전문
```

## 의존성

### 필수 (자동 설치)

| 패키지 | 라이선스 | 용도 |
|--------|----------|------|
| olefile | BSD | OLE2 파일 파싱 |
| pyyaml | MIT | YAML 출력 |
| tqdm | MIT | 진행률 표시 |

### HWP 3.x 변환 (선택)

| 패키지/도구 | 라이선스 | 용도 |
|-------------|----------|------|
| LibreOffice | MPL-2.0 | HWP → PDF 변환 |
| docling | MIT | PDF 구조 추출 (권장) |
| poppler-utils | GPL-2.0 | pdftotext (Docling 폴백) |

## 에러 처리

```python
from hwp2yaml import extract_hwp_text, HWPVersion, detect_hwp_version

# 파일이 존재하지 않는 경우
result = extract_hwp_text("nonexistent.hwp")
if not result.success:
    print(f"에러: {result.error}")  # "파일을 찾을 수 없습니다"

# 지원하지 않는 형식
version = detect_hwp_version("document.pdf")
if version == HWPVersion.UNKNOWN:
    print("HWP 파일이 아닙니다")

# 암호화된 문서
result = extract_hwp_text("encrypted.hwp")
if not result.success:
    print(f"에러: {result.error}")  # "암호화된 문서입니다"
```

## 제한사항

| 제한 | 설명 |
|------|------|
| 암호화된 문서 | 미지원 (HWP 파일 형식 한계) |
| 이미지/OLE 객체 | 미지원 (텍스트 추출에 집중) |
| HWP 3.x | LibreOffice 설치 필요 |
| DRM 보호 문서 | 미지원 |

## Changelog

### v0.6.1 (2026-01-19)

**버그 수정:**
1. **테이블 종료 감지 개선** - 레코드 레벨 기반 상태 추적으로 테이블 후 단락이 셀로 잘못 포함되는 문제 해결
2. **다중 레코드 단락 누적** - 여러 PARA_TEXT 레코드로 분할된 긴 단락을 완전히 캡처
3. **셀 진행 로직 개선** - 테이블 외부에서 LIST_HEADER 오작동 방지, 범위 초과 크래시 방지
4. **HWPX 섹션 순서 정렬** - 숫자 기준 정렬로 section10.xml이 section2.xml 앞에 오는 문제 해결
5. **네임스페이스 안전 파싱** - 네임스페이스 포함 HWPX XML 파싱 개선, Fallback 메커니즘 추가
6. **마지막 셀 텍스트 누출 방지** - 테이블 종료 시 대기 중인 텍스트 flush 순서 수정

**테스트:**
- 420+ 줄 신규 테스트 코드
- 26개 테스트 케이스 (구조 보존, HWPX 파싱, 에지 케이스)

### v0.6.0 (2026-01-19)

- HWP 5.x 구조 보존 추출 기능 추가
- 단락, 테이블, 섹션 구조 완전 보존
- YAML 스키마 정의

### v0.5.0 (2025-12)

- HWP → YAML 직접 변환
- HWP 5.x 직접 파싱
- HWPX 지원 추가

## 라이선스

MIT License

Copyright (c) 2025-2026 Dyarchy Project

## 기여

이슈와 PR을 환영합니다:
- GitHub: https://github.com/seunghyuoffice-design/hwp2yaml
