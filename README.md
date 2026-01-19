# hwp-parser

HWP 5.x 문서에서 텍스트를 추출하는 Python 라이브러리.

## 특징

- **순수 Python**: pyhwp 등 AGPL 라이브러리 대체
- **MIT 라이선스**: 상업적 사용 가능
- **두 가지 추출 전략**: PrvText (빠름) + BodyText (완전)
- **배치 처리**: 멀티프로세싱 병렬 처리
- **메타데이터 보존**: HWP 버전, 압축 여부 등

## 설치

```bash
pip install hwp-parser
```

또는 소스에서:

```bash
git clone https://github.com/yourusername/hwp-parser.git
cd hwp-parser
pip install -e .
```

## 사용법

### 단일 파일 추출

```python
from hwp_parser import extract_hwp_text

result = extract_hwp_text("document.hwp")

if result.success:
    print(result.text)
    print(f"방법: {result.method}")  # "prvtext" 또는 "bodytext"
    print(f"글자수: {result.char_count}")
```

### 배치 처리

```python
from hwp_parser import BatchProcessor

processor = BatchProcessor(workers=4)
result = processor.process_directory("/path/to/hwp/files")

print(f"성공: {result.success}/{result.total}")
```

### CLI

```bash
# 단일 파일
hwp-parser extract document.hwp

# 배치 처리
hwp-parser batch /path/to/files -o output/ --format jsonl

# 파일 정보
hwp-parser info document.hwp
```

## 의존성

- `olefile` (BSD) - OLE2 파일 파싱
- `pyyaml` (MIT) - YAML 출력
- `tqdm` (MIT) - 진행률 표시

## 라이선스

MIT License

## 제한사항

- HWP 5.x만 지원 (3.x, 4.x 미지원)
- 암호화된 문서 미지원
- HWPX (XML 기반) 미지원
