"""YAML 출력 모듈"""

import os
import json
from pathlib import Path
from datetime import datetime
from typing import Callable

import yaml

from .models import ExtractResult, BatchResult, TrainingData


class YAMLExporter:
    """
    Qwen3 학습용 YAML 변환기

    메타데이터 보존:
    - 원본 파일 경로
    - 분류 (disputes/materials)
    - 크롤링 메타데이터 (article_id, title, date 등)
    - HWP 메타데이터 (버전, 압축 여부 등)
    - 추출 방법 (prvtext/bodytext)
    """

    def __init__(
        self,
        output_dir: str,
        category_detector: Callable[[str], str] | None = None,
    ):
        """
        Args:
            output_dir: 출력 디렉토리
            category_detector: 파일 경로 → 카테고리 함수
        """
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

        self.category_detector = category_detector or self._default_category

    @staticmethod
    def _default_category(filepath: str) -> str:
        """기본 카테고리 감지 (경로 기반)"""
        path_lower = filepath.lower()
        if "disputes" in path_lower or "분쟁" in path_lower:
            return "disputes"
        elif "materials" in path_lower or "보도" in path_lower:
            return "materials"
        else:
            return "unknown"

    def result_to_training_data(
        self,
        result: ExtractResult,
        external_metadata: dict | None = None,
    ) -> TrainingData | None:
        """
        ExtractResult를 TrainingData로 변환

        Args:
            result: 추출 결과
            external_metadata: 외부 메타데이터 (크롤링 데이터)

        Returns:
            TrainingData 또는 None (실패 시)
        """
        if not result.success or not result.text:
            return None

        # 카테고리 감지
        category = self.category_detector(result.filepath)

        # 제목 추출 (우선순위: 외부 메타데이터 > 파일명 > 본문 첫 줄)
        title = self._extract_title(result, external_metadata)

        # 메타데이터 병합
        metadata = self._merge_metadata(result, external_metadata)

        return TrainingData(
            source=result.filepath,
            category=category,
            title=title,
            content=result.text,
            metadata=metadata,
        )

    def _extract_title(
        self,
        result: ExtractResult,
        external_metadata: dict | None,
    ) -> str:
        """제목 추출"""
        # 1. 외부 메타데이터에서 제목
        if external_metadata and "title" in external_metadata:
            return external_metadata["title"]

        # 2. 파일명에서 추출
        filename = os.path.basename(result.filepath)
        name_without_ext = os.path.splitext(filename)[0]

        # 괄호 제거, 언더스코어 공백 변환
        title = name_without_ext.replace("_", " ").strip()

        # 너무 짧으면 본문 첫 줄 사용
        if len(title) < 5 and result.text:
            first_line = result.text.split("\n")[0].strip()
            if first_line:
                return first_line[:100]

        return title

    def _merge_metadata(
        self,
        result: ExtractResult,
        external_metadata: dict | None,
    ) -> dict:
        """메타데이터 병합"""
        metadata = {}

        # HWP 메타데이터
        if result.metadata:
            metadata["hwp"] = {
                "version": str(result.metadata.version) if result.metadata.version else None,
                "compressed": result.metadata.is_compressed,
                "file_size_bytes": result.metadata.file_size_bytes,
            }

        # 추출 정보
        metadata["extraction"] = {
            "method": result.method,
            "char_count": result.char_count,
            "extracted_at": result.extracted_at.isoformat(),
        }

        # 외부 메타데이터 (크롤링)
        if external_metadata:
            metadata["crawl"] = {
                k: v for k, v in external_metadata.items()
                if k not in ("content", "text", "body")  # 본문 제외
            }

        return metadata

    def export_single(
        self,
        result: ExtractResult,
        external_metadata: dict | None = None,
    ) -> str | None:
        """
        단일 결과 YAML 저장

        Returns:
            저장된 파일 경로 또는 None
        """
        training_data = self.result_to_training_data(result, external_metadata)
        if not training_data:
            return None

        # 파일명 생성 (원본 파일명 기반)
        base_name = Path(result.filepath).stem
        output_path = self.output_dir / f"{base_name}.yaml"

        with open(output_path, "w", encoding="utf-8") as f:
            yaml.dump(
                training_data.to_dict(),
                f,
                allow_unicode=True,
                default_flow_style=False,
                sort_keys=False,
            )

        return str(output_path)

    def export_batch(
        self,
        batch_result: BatchResult,
        metadata_getter: Callable[[str], dict] | None = None,
    ) -> list[str]:
        """
        배치 결과 YAML 저장

        Args:
            batch_result: 배치 처리 결과
            metadata_getter: 파일 경로 → 외부 메타데이터 함수

        Returns:
            저장된 파일 경로 목록
        """
        saved_paths = []

        for result in batch_result.results:
            if not result.success:
                continue

            external_meta = None
            if metadata_getter:
                external_meta = metadata_getter(result.filepath)

            path = self.export_single(result, external_meta)
            if path:
                saved_paths.append(path)

        return saved_paths

    def export_batch_jsonl(
        self,
        batch_result: BatchResult,
        output_file: str,
        metadata_getter: Callable[[str], dict] | None = None,
    ) -> int:
        """
        배치 결과 JSONL 저장 (대용량 처리용)

        Args:
            batch_result: 배치 처리 결과
            output_file: 출력 파일 경로
            metadata_getter: 파일 경로 → 외부 메타데이터 함수

        Returns:
            저장된 레코드 수
        """
        count = 0

        with open(output_file, "w", encoding="utf-8") as f:
            for result in batch_result.results:
                if not result.success:
                    continue

                external_meta = None
                if metadata_getter:
                    external_meta = metadata_getter(result.filepath)

                training_data = self.result_to_training_data(result, external_meta)
                if training_data:
                    f.write(json.dumps(training_data.to_dict(), ensure_ascii=False))
                    f.write("\n")
                    count += 1

        return count

    def export_failed_log(
        self,
        batch_result: BatchResult,
        output_file: str,
    ) -> int:
        """
        실패 로그 JSONL 저장

        Args:
            batch_result: 배치 처리 결과
            output_file: 출력 파일 경로

        Returns:
            저장된 레코드 수
        """
        count = 0

        with open(output_file, "w", encoding="utf-8") as f:
            for result in batch_result.results:
                if result.success:
                    continue

                log_entry = {
                    "filepath": result.filepath,
                    "error": result.error,
                    "method": result.method,
                    "timestamp": result.extracted_at.isoformat(),
                }
                f.write(json.dumps(log_entry, ensure_ascii=False))
                f.write("\n")
                count += 1

        return count
