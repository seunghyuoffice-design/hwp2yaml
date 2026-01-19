# hwp-parser

HWP 5.x / HWPX / HWP 3.x 문서에서 텍스트를 추출하는 Python 라이브러리.

## 지원 형식

| 형식 | 지원 | 방식 | 비고 |
|------|------|------|------|
| HWP 5.x | ✅ | 직접 파싱 | OLE2 기반 (2002~) |
| HWPX | ✅ | 직접 파싱 | XML + ZIP 기반 (2014~) |
| HWP 3.x | ✅ | LibreOffice + Docling | 구형식 (1990년대) |

## 특징

- **순수 Python**: pyhwp 등 AGPL 라이브러리 대체
- **MIT 라이선스**: 상업적 사용 가능
- **전 버전 지원**: HWP 3.x ~ HWPX 모두 처리
- **트리아지**: 파일 버전 자동 감지 및 분류
- **표 변환**: 표 태그를 마크다운으로 변환
- **배치 처리**: 멀티프로세싱 병렬 처리

## 설치

```bash
pip install hwp-parser
```

또는 소스에서:

```bash
git clone https://github.com/seunghyuoffice-design/hwp-parser.git
cd hwp-parser
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
from hwp_parser import detect_hwp_version, HWPVersion, triage_directory

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
from hwp_parser import extract_hwp_text

result = extract_hwp_text("document.hwp")

if result.success:
    print(result.text)
    print(f"방법: {result.method}")  # "prvtext", "bodytext", "hwpx"
```

### HWP 3.x 변환 (LibreOffice + Docling)

```python
from hwp_parser import convert_hwp3, batch_convert_hwp3

# 단일 파일 변환
result = convert_hwp3("old_document.hwp")

if result.success:
    print(result.text)          # 추출된 텍스트
    print(result.markdown)      # 마크다운 (Docling 사용 시)
    print(f"테이블: {len(result.tables)}")  # 추출된 테이블
    print(f"방법: {result.method}")  # "hwp3_docling" 또는 "hwp3_pdftotext"

# 배치 변환
results = batch_convert_hwp3(
    filepaths=["file1.hwp", "file2.hwp"],
    keep_pdf=True,  # PDF 유지
    pdf_output_dir="/path/to/pdfs",
)

for r in results:
    print(f"{r.filepath}: {'성공' if r.success else r.error}")
```

### 배치 처리 (HWP 5.x/HWPX)

```python
from hwp_parser import BatchProcessor

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
