# hwp2yaml

HWP 5.x / HWPX / HWP 3.x 문서를 YAML로 직접 변환하는 Python 라이브러리.

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

```bash
pip install hwp2yaml
```

또는 소스에서:

```bash
git clone https://github.com/seunghyuoffice-design/hwp2yaml.git
cd hwp2yaml
pip install -e .
```

### HWP 3.x 변환 의존성 (선택)

HWP 3.x 변환을 사용하려면 추가 설치 필요:

```bash
# LibreOffice (필수)
sudo apt install libreoffice

# Docling (권장 - 구조 추출)
pip install docling

# pdftotext (Docling 없을 때 폴백)
sudo apt install poppler-utils
```

## 사용법

### 트리아지 (버전 확인)

```python
from hwp2yaml import detect_hwp_version, HWPVersion, triage_directory

# 단일 파일 버전 확인
version = detect_hwp_version("document.hwp")
if version == HWPVersion.HWP_3X:
    print("HWP 3.x - LibreOffice 변환 필요")
elif version == HWPVersion.HWP_5X:
    print("HWP 5.x - 직접 처리 가능")
elif version == HWPVersion.HWPX:
    print("HWPX - 직접 처리 가능")

# 디렉토리 트리아지
summary = triage_directory("/path/to/files")
print(f"총: {summary.total}")
print(f"HWP 5.x: {summary.hwp5_count}")
print(f"HWPX: {summary.hwpx_count}")
print(f"HWP 3.x: {summary.hwp3_count}")
```

### HWP 5.x / HWPX 텍스트 추출

```python
from hwp2yaml import extract_hwp_text

result = extract_hwp_text("document.hwp")

if result.success:
    print(result.text)
    print(f"방법: {result.method}")  # "prvtext", "bodytext", "hwpx"
```

### HWP 5.x 구조 보존 추출 (NEW)

단락, 테이블, 섹션 구조를 보존하여 YAML로 변환:

```python
from hwp2yaml import extract_hwp_structure

result = extract_hwp_structure("document.hwp")

if result.success:
    print(f"방법: {result.method}")  # "hwp5_structure"

    # 구조 정보
    for section in result.structure["sections"]:
        print(f"단락 수: {len(section['paragraphs'])}")
        print(f"테이블 수: {len(section['tables'])}")

    # 테이블 추출
    for table in result.tables:
        print(f"테이블: {table['rows']}x{table['cols']}")
        print(table['data'])  # 2D 배열

    # YAML 출력
    print(result.to_yaml())
```

### 구조 보존 YAML 스키마

```yaml
metadata:
  source_file: /path/to/document.hwp
  method: hwp5_structure
  extracted_at: "2026-01-19T..."

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

raw_text: 평탄화된 텍스트
```

### HWP 3.x → YAML 변환

```python
from hwp2yaml import convert_hwp3, convert_to_yaml, batch_convert_to_yaml

# 단일 파일 변환
result = convert_hwp3("old_document.hwp")

if result.success:
    print(result.text)           # 추출된 텍스트
    print(f"테이블: {len(result.tables)}")  # 추출된 테이블
    print(f"방법: {result.method}")  # "hwp3_docling" 또는 "hwp3_pdftotext"

# YAML 딕셔너리로 변환
yaml_dict = result.to_yaml_dict()

# YAML 문자열로 변환
yaml_str = result.to_yaml()

# 직접 YAML 파일로 저장
yaml_dict = convert_to_yaml("old_document.hwp", output_path="output.yaml")
```

### 배치 YAML 변환

```python
from hwp2yaml import batch_convert_to_yaml

# 배치 변환 (개별 YAML 파일 + 통합 파일)
yaml_dicts = batch_convert_to_yaml(
    filepaths=["file1.hwp", "file2.hwp"],
    output_dir="/path/to/yaml",           # 개별 파일 저장
    combined_output="/path/to/all.yaml",  # 통합 파일 저장
)
```

### YAML 출력 스키마

```yaml
metadata:
  case_id: "2002-4"
  source: fss_disputes
  source_file: /path/to/file.hwp
  version: hwp3
  method: hwp3_docling
  converted_at: "2026-01-19T..."

content:
  parties: 당사자 정보
  request: 신청취지
  facts: 사실관계
  applicant_claim: 신청인 주장
  respondent_claim: 피신청인 주장
  decision: 위원회 판단

tables:
  - table_id: table_0
    rows: 3
    cols: 2
    cells:
      - {row: 0, col: 0, text: "..."}

raw_text: 원본 전문
```

### 배치 처리 (HWP 5.x/HWPX)

```python
from hwp2yaml import BatchProcessor

processor = BatchProcessor(workers=4)
result = processor.process_directory("/path/to/files")

print(f"성공: {result.success}/{result.total}")
```

## 의존성

### 필수
- `olefile` (BSD) - OLE2 파일 파싱
- `pyyaml` (MIT) - YAML 출력
- `tqdm` (MIT) - 진행률 표시

### HWP 3.x 변환 (선택)
- LibreOffice (MPL-2.0) - HWP → PDF 변환
- `docling` (MIT) - PDF 구조 추출 (권장)
- poppler-utils - pdftotext (Docling 폴백)

## 라이선스

MIT License

## 제한사항

- 암호화된 문서 미지원
- 이미지/OLE 객체 추출 미지원
- HWP 3.x 변환 시 LibreOffice 설치 필요
