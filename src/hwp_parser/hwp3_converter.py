"""HWP 3.x 변환기 - LibreOffice + Docling 기반

HWP 3.x는 1990년대 구형 바이너리 포맷으로 직접 파싱이 어려움.
LibreOffice로 PDF 변환 후 Docling으로 구조 추출.

흐름:
    HWP 3.x → LibreOffice → PDF → Docling → 구조화된 출력
"""

import os
import tempfile
import subprocess
import shutil
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, List, Dict, Any

from .triage import detect_hwp_version, HWPVersion


@dataclass
class ConversionResult:
    """HWP 3.x 변환 결과"""
    filepath: str
    success: bool
    text: Optional[str] = None
    tables: List[Dict[str, Any]] = field(default_factory=list)
    markdown: Optional[str] = None
    method: str = "hwp3_docling"
    error: Optional[str] = None
    pdf_path: Optional[str] = None


class HWP3Converter:
    """
    HWP 3.x 변환기

    LibreOffice를 사용하여 PDF로 변환 후
    Docling으로 구조화된 텍스트 추출
    """

    def __init__(
        self,
        libreoffice_path: str = "libreoffice",
        keep_pdf: bool = False,
        pdf_output_dir: Optional[str] = None,
    ):
        """
        Args:
            libreoffice_path: LibreOffice 실행 경로
            keep_pdf: 변환된 PDF 유지 여부
            pdf_output_dir: PDF 저장 디렉토리 (None이면 임시 디렉토리)
        """
        self.libreoffice_path = libreoffice_path
        self.keep_pdf = keep_pdf
        self.pdf_output_dir = pdf_output_dir

        # LibreOffice 확인
        self._check_libreoffice()

        # Docling 확인
        self._docling_available = self._check_docling()

    def _check_libreoffice(self) -> None:
        """LibreOffice 설치 확인"""
        try:
            result = subprocess.run(
                [self.libreoffice_path, "--version"],
                capture_output=True,
                text=True,
                timeout=10,
            )
            if result.returncode != 0:
                raise RuntimeError("LibreOffice 실행 실패")
        except FileNotFoundError:
            raise RuntimeError(
                f"LibreOffice를 찾을 수 없습니다: {self.libreoffice_path}\n"
                "설치: sudo apt install libreoffice"
            )

    def _check_docling(self) -> bool:
        """Docling 설치 확인"""
        try:
            from docling.document_converter import DocumentConverter
            return True
        except ImportError:
            return False

    def convert(self, filepath: str) -> ConversionResult:
        """
        HWP 3.x 파일을 변환

        Args:
            filepath: HWP 3.x 파일 경로

        Returns:
            ConversionResult
        """
        filepath = str(filepath)

        # 버전 확인
        version = detect_hwp_version(filepath)
        if version != HWPVersion.HWP_3X:
            return ConversionResult(
                filepath=filepath,
                success=False,
                error=f"HWP 3.x가 아님: {version.value}",
            )

        # PDF 변환
        try:
            pdf_path = self._convert_to_pdf(filepath)
        except Exception as e:
            return ConversionResult(
                filepath=filepath,
                success=False,
                error=f"PDF 변환 실패: {e}",
            )

        # Docling으로 추출
        if self._docling_available:
            try:
                result = self._extract_with_docling(filepath, pdf_path)
                return result
            except Exception as e:
                # Docling 실패 시 pdftotext 폴백
                return self._extract_with_pdftotext(filepath, pdf_path)
        else:
            # Docling 없으면 pdftotext 사용
            return self._extract_with_pdftotext(filepath, pdf_path)

    def _convert_to_pdf(self, filepath: str) -> str:
        """LibreOffice로 PDF 변환"""
        if self.pdf_output_dir:
            output_dir = self.pdf_output_dir
            os.makedirs(output_dir, exist_ok=True)
        else:
            output_dir = tempfile.mkdtemp(prefix="hwp3_")

        # LibreOffice 변환 실행
        result = subprocess.run(
            [
                self.libreoffice_path,
                "--headless",
                "--convert-to", "pdf",
                "--outdir", output_dir,
                filepath,
            ],
            capture_output=True,
            text=True,
            timeout=120,
        )

        if result.returncode != 0:
            raise RuntimeError(f"LibreOffice 오류: {result.stderr}")

        # PDF 파일 경로 찾기
        basename = os.path.splitext(os.path.basename(filepath))[0]
        pdf_path = os.path.join(output_dir, f"{basename}.pdf")

        if not os.path.exists(pdf_path):
            raise RuntimeError(f"PDF 파일 생성 실패: {pdf_path}")

        return pdf_path

    def _extract_with_docling(self, filepath: str, pdf_path: str) -> ConversionResult:
        """Docling으로 구조 추출"""
        from docling.document_converter import DocumentConverter

        converter = DocumentConverter()
        result = converter.convert(pdf_path)

        # 마크다운 추출
        markdown = result.document.export_to_markdown()

        # 테이블 추출
        tables = []
        for i, table in enumerate(result.document.tables):
            table_data = {
                "table_id": f"table_{i}",
                "rows": table.num_rows,
                "cols": table.num_cols,
                "cells": [],
            }
            # 셀 데이터 추출
            try:
                for row_idx, row in enumerate(table.data):
                    for col_idx, cell in enumerate(row):
                        table_data["cells"].append({
                            "row": row_idx,
                            "col": col_idx,
                            "text": str(cell) if cell else "",
                        })
            except Exception:
                pass
            tables.append(table_data)

        # 텍스트 추출 (마크다운에서)
        text = markdown

        # PDF 정리
        if not self.keep_pdf:
            self._cleanup_pdf(pdf_path)

        return ConversionResult(
            filepath=filepath,
            success=True,
            text=text,
            tables=tables,
            markdown=markdown,
            method="hwp3_docling",
            pdf_path=pdf_path if self.keep_pdf else None,
        )

    def _extract_with_pdftotext(self, filepath: str, pdf_path: str) -> ConversionResult:
        """pdftotext로 텍스트 추출 (Docling 폴백)"""
        try:
            result = subprocess.run(
                ["pdftotext", "-layout", pdf_path, "-"],
                capture_output=True,
                text=True,
                timeout=60,
            )

            if result.returncode != 0:
                raise RuntimeError(f"pdftotext 오류: {result.stderr}")

            text = result.stdout.strip()

            # PDF 정리
            if not self.keep_pdf:
                self._cleanup_pdf(pdf_path)

            return ConversionResult(
                filepath=filepath,
                success=True,
                text=text,
                tables=[],
                markdown=None,
                method="hwp3_pdftotext",
                pdf_path=pdf_path if self.keep_pdf else None,
            )

        except FileNotFoundError:
            return ConversionResult(
                filepath=filepath,
                success=False,
                error="pdftotext를 찾을 수 없습니다. poppler-utils 설치 필요",
            )
        except Exception as e:
            return ConversionResult(
                filepath=filepath,
                success=False,
                error=f"pdftotext 실패: {e}",
            )

    def _cleanup_pdf(self, pdf_path: str) -> None:
        """PDF 파일 정리"""
        try:
            if os.path.exists(pdf_path):
                os.remove(pdf_path)
            # 임시 디렉토리 정리
            parent = os.path.dirname(pdf_path)
            if parent.startswith(tempfile.gettempdir()) and os.path.isdir(parent):
                shutil.rmtree(parent, ignore_errors=True)
        except Exception:
            pass


def convert_hwp3(
    filepath: str,
    keep_pdf: bool = False,
    pdf_output_dir: Optional[str] = None,
) -> ConversionResult:
    """
    HWP 3.x 파일 변환 (편의 함수)

    Args:
        filepath: HWP 3.x 파일 경로
        keep_pdf: 변환된 PDF 유지 여부
        pdf_output_dir: PDF 저장 디렉토리

    Returns:
        ConversionResult
    """
    converter = HWP3Converter(
        keep_pdf=keep_pdf,
        pdf_output_dir=pdf_output_dir,
    )
    return converter.convert(filepath)


def batch_convert_hwp3(
    filepaths: List[str],
    keep_pdf: bool = False,
    pdf_output_dir: Optional[str] = None,
    progress: bool = True,
) -> List[ConversionResult]:
    """
    HWP 3.x 파일 배치 변환

    Args:
        filepaths: HWP 3.x 파일 경로 목록
        keep_pdf: 변환된 PDF 유지 여부
        pdf_output_dir: PDF 저장 디렉토리
        progress: 진행률 표시

    Returns:
        ConversionResult 목록
    """
    converter = HWP3Converter(
        keep_pdf=keep_pdf,
        pdf_output_dir=pdf_output_dir,
    )

    results = []
    iterator = filepaths

    if progress:
        try:
            from tqdm import tqdm
            iterator = tqdm(filepaths, desc="HWP 3.x 변환")
        except ImportError:
            pass

    for filepath in iterator:
        result = converter.convert(filepath)
        results.append(result)

    return results


__all__ = [
    "HWP3Converter",
    "ConversionResult",
    "convert_hwp3",
    "batch_convert_hwp3",
]
