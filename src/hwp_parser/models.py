"""HWP 파서 데이터 모델"""

from dataclasses import dataclass, field
from typing import Literal
from datetime import datetime


@dataclass
class RecordHeader:
    """HWP 레코드 헤더"""
    tag_id: int
    level: int
    size: int

    @property
    def tag_name(self) -> str:
        """태그 ID를 사람이 읽을 수 있는 이름으로 변환"""
        from .constants import (
            HWPTAG_PARA_TEXT, HWPTAG_PARA_HEADER, HWPTAG_PARA_CHAR_SHAPE,
            HWPTAG_TABLE, HWPTAG_CTRL_HEADER
        )
        names = {
            HWPTAG_PARA_TEXT: "PARA_TEXT",
            HWPTAG_PARA_HEADER: "PARA_HEADER",
            HWPTAG_PARA_CHAR_SHAPE: "PARA_CHAR_SHAPE",
            HWPTAG_TABLE: "TABLE",
            HWPTAG_CTRL_HEADER: "CTRL_HEADER",
        }
        return names.get(self.tag_id, f"TAG_{self.tag_id:04X}")


@dataclass
class HWPVersion:
    """HWP 버전 정보"""
    major: int
    minor: int
    build: int
    revision: int

    def __str__(self) -> str:
        return f"{self.major}.{self.minor}.{self.build}.{self.revision}"

    @property
    def is_5x(self) -> bool:
        """HWP 5.x 버전인지 확인"""
        return self.major == 5


@dataclass
class HWPMetadata:
    """HWP 문서 메타데이터"""
    filepath: str
    filename: str
    version: HWPVersion | None = None
    is_compressed: bool = False
    is_encrypted: bool = False
    file_size_bytes: int = 0
    streams: list[str] = field(default_factory=list)


@dataclass
class ExtractResult:
    """텍스트 추출 결과"""
    filepath: str
    success: bool
    text: str | None
    method: Literal["prvtext", "bodytext", "failed"]
    error: str | None = None
    metadata: HWPMetadata | None = None
    char_count: int = 0
    extracted_at: datetime = field(default_factory=datetime.now)

    def __post_init__(self):
        if self.text:
            self.char_count = len(self.text)


@dataclass
class BatchResult:
    """배치 처리 결과"""
    total: int
    success: int
    failed: int
    results: list[ExtractResult] = field(default_factory=list)
    started_at: datetime = field(default_factory=datetime.now)
    finished_at: datetime | None = None

    @property
    def success_rate(self) -> float:
        """성공률 (0.0 ~ 1.0)"""
        if self.total == 0:
            return 0.0
        return self.success / self.total

    @property
    def failed_files(self) -> list[str]:
        """실패한 파일 목록"""
        return [r.filepath for r in self.results if not r.success]

    def to_summary(self) -> dict:
        """요약 정보 반환"""
        return {
            "total": self.total,
            "success": self.success,
            "failed": self.failed,
            "success_rate": f"{self.success_rate:.1%}",
            "started_at": self.started_at.isoformat(),
            "finished_at": self.finished_at.isoformat() if self.finished_at else None,
        }


@dataclass
class TrainingData:
    """Qwen3 학습용 데이터"""
    source: str              # 원본 파일 경로
    category: str            # 분류 (disputes, materials 등)
    title: str               # 문서 제목
    content: str             # 본문 텍스트
    metadata: dict = field(default_factory=dict)

    def to_dict(self) -> dict:
        """YAML 출력용 딕셔너리"""
        return {
            "source": self.source,
            "category": self.category,
            "title": self.title,
            "content": self.content,
            "metadata": self.metadata,
        }
