# hwp-parser

HWP 5.x / HWPX 문서에서 텍스트를 추출하는 Python 라이브러리.

## 지원 형식

| 형식 | 지원 | 비고 |
|------|------|------|
| HWP 5.x | ✅ | OLE2 기반 (2002~) |
| HWPX | ✅ | XML + ZIP 기반 (2014~) |
| HWP 3.x | ❌ | **미지원** (1990년대 구형식) |

> ⚠️ **HWP 3.x는 지원하지 않습니다.** 트리아지 기능으로 버전을 확인한 후 처리하세요.

## 특징

- **순수 Python**: pyhwp 등 AGPL 라이브러리 대체
- **MIT 라이선스**: 상업적 사용 가능
- **HWPX 지원**: XML 기반 신형식 처리
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

## 사용법

### 트리아지 (버전 확인)

```python
from hwp_parser import detect_hwp_version, HWPVersion, triage_directory

# 단일 파일 버전 확인
version = detect_hwp_version("document.hwp")
if version == HWPVersion.HWP_3X:
    print("HWP 3.x - 미지원")
elif version == HWPVersion.HWP_5X:
    print("HWP 5.x - 처리 가능")
elif version == HWPVersion.HWPX:
    print("HWPX - 처리 가능")

# 디렉토리 트리아지
summary = triage_directory("/path/to/files")
print(f"총: {summary.total}")
print(f"HWP 5.x: {summary.hwp5_count}")
print(f"HWPX: {summary.hwpx_count}")
print(f"HWP 3.x (스킵): {summary.hwp3_count}")
```

### 텍스트 추출

```python
from hwp_parser import extract_hwp_text

result = extract_hwp_text("document.hwp")

if result.success:
    print(result.text)
    print(f"방법: {result.method}")  # "prvtext", "bodytext", "hwpx"
```

### 배치 처리

```python
from hwp_parser import BatchProcessor

processor = BatchProcessor(workers=4)
result = processor.process_directory("/path/to/files")

print(f"성공: {result.success}/{result.total}")
```

## 의존성

- `olefile` (BSD) - OLE2 파일 파싱
- `pyyaml` (MIT) - YAML 출력
- `tqdm` (MIT) - 진행률 표시

## 라이선스

MIT License

## 제한사항

- **HWP 3.x 미지원** - 1990년대 구형식, 별도 변환 필요
- 암호화된 문서 미지원
- 이미지/OLE 객체 추출 미지원
