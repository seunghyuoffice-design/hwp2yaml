"""배치 처리 모듈"""

import os
import json
import signal
from pathlib import Path
from datetime import datetime
from concurrent.futures import ProcessPoolExecutor, as_completed, TimeoutError
from typing import Callable
from multiprocessing import cpu_count

from tqdm import tqdm

from .extractor import extract_hwp_text
from .models import ExtractResult, BatchResult
from .constants import TIMEOUT_SECONDS


def _worker_extract(filepath: str) -> ExtractResult:
    """워커 프로세스에서 실행되는 추출 함수"""
    return extract_hwp_text(filepath)


class BatchProcessor:
    """
    HWP 파일 배치 처리기

    병렬 처리로 다수의 HWP 파일에서 텍스트 추출
    """

    def __init__(
        self,
        workers: int | None = None,
        timeout: int = TIMEOUT_SECONDS,
        metadata_mapper: Callable[[str], dict] | None = None,
    ):
        """
        Args:
            workers: 워커 프로세스 수 (기본: CPU 코어의 50%)
            timeout: 파일당 타임아웃 (초)
            metadata_mapper: 파일 경로 → 외부 메타데이터 매핑 함수
        """
        if workers is None:
            workers = max(1, cpu_count() // 2)
        self.workers = workers
        self.timeout = timeout
        self.metadata_mapper = metadata_mapper

    def process_files(
        self,
        files: list[str],
        progress: bool = True,
    ) -> BatchResult:
        """
        파일 목록 처리

        Args:
            files: HWP 파일 경로 목록
            progress: 진행률 표시 여부

        Returns:
            BatchResult 객체
        """
        total = len(files)
        success = 0
        failed = 0
        results: list[ExtractResult] = []

        started_at = datetime.now()

        with ProcessPoolExecutor(max_workers=self.workers) as executor:
            # 작업 제출
            future_to_file = {
                executor.submit(_worker_extract, f): f for f in files
            }

            # 진행률 표시
            iterator = as_completed(future_to_file)
            if progress:
                iterator = tqdm(iterator, total=total, desc="HWP 추출")

            for future in iterator:
                filepath = future_to_file[future]

                try:
                    result = future.result(timeout=self.timeout)

                    # 외부 메타데이터 병합
                    if self.metadata_mapper and result.metadata:
                        external_meta = self.metadata_mapper(filepath)
                        if external_meta and result.metadata:
                            # HWPMetadata에 추가 정보 저장 (확장)
                            if not hasattr(result.metadata, 'external'):
                                object.__setattr__(result.metadata, 'external', external_meta)

                    results.append(result)

                    if result.success:
                        success += 1
                    else:
                        failed += 1

                except TimeoutError:
                    failed += 1
                    results.append(ExtractResult(
                        filepath=filepath,
                        success=False,
                        text=None,
                        method="failed",
                        error=f"타임아웃 ({self.timeout}초)",
                    ))

                except Exception as e:
                    failed += 1
                    results.append(ExtractResult(
                        filepath=filepath,
                        success=False,
                        text=None,
                        method="failed",
                        error=str(e),
                    ))

        finished_at = datetime.now()

        return BatchResult(
            total=total,
            success=success,
            failed=failed,
            results=results,
            started_at=started_at,
            finished_at=finished_at,
        )

    def process_directory(
        self,
        directory: str,
        recursive: bool = True,
        progress: bool = True,
    ) -> BatchResult:
        """
        디렉토리 내 모든 HWP 파일 처리

        Args:
            directory: 디렉토리 경로
            recursive: 하위 디렉토리 포함 여부
            progress: 진행률 표시 여부

        Returns:
            BatchResult 객체
        """
        path = Path(directory)

        if recursive:
            files = list(path.rglob("*.hwp"))
        else:
            files = list(path.glob("*.hwp"))

        # 문자열 경로로 변환
        file_paths = [str(f) for f in files]

        return self.process_files(file_paths, progress=progress)


class MetadataMapper:
    """
    외부 메타데이터 매퍼

    크롤링 시 수집한 메타데이터를 HWP 파일과 매핑
    """

    def __init__(self, metadata_file: str | None = None):
        """
        Args:
            metadata_file: 메타데이터 JSONL 파일 경로
        """
        self._mapping: dict[str, dict] = {}

        if metadata_file:
            self.load(metadata_file)

    def load(self, filepath: str) -> None:
        """JSONL 파일에서 메타데이터 로드"""
        if not os.path.isfile(filepath):
            return

        with open(filepath, "r", encoding="utf-8") as f:
            for line in f:
                try:
                    data = json.loads(line.strip())
                    # article_id 또는 파일명으로 키 생성
                    key = self._make_key(data)
                    if key:
                        self._mapping[key] = data
                except json.JSONDecodeError:
                    continue

    def _make_key(self, data: dict) -> str | None:
        """메타데이터에서 키 생성"""
        if "article_id" in data:
            return str(data["article_id"])
        if "filename" in data:
            return data["filename"]
        return None

    def get(self, filepath: str) -> dict:
        """
        파일 경로에서 메타데이터 조회

        Args:
            filepath: HWP 파일 경로

        Returns:
            메타데이터 딕셔너리 (없으면 빈 딕셔너리)
        """
        filename = os.path.basename(filepath)

        # 파일명에서 article_id 추출 시도
        # 예: "133695_0.hwp" → "133695"
        base = filename.split("_")[0].split(".")[0]

        if base in self._mapping:
            return self._mapping[base]

        if filename in self._mapping:
            return self._mapping[filename]

        return {}

    def __call__(self, filepath: str) -> dict:
        """함수처럼 호출 가능"""
        return self.get(filepath)
