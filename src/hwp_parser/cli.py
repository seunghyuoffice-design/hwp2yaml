#!/usr/bin/env python3
"""HWP 파서 CLI"""

import os
import sys
import json
import argparse
from pathlib import Path

from .extractor import extract_hwp_text
from .batch import BatchProcessor, MetadataMapper
from .exporter import YAMLExporter


def cmd_extract(args):
    """단일 파일 추출"""
    result = extract_hwp_text(args.file)

    if result.success:
        if args.output:
            with open(args.output, "w", encoding="utf-8") as f:
                f.write(result.text)
            print(f"✓ 저장됨: {args.output}")
        else:
            print(result.text)
        return 0
    else:
        print(f"✗ 추출 실패: {result.error}", file=sys.stderr)
        return 1


def cmd_batch(args):
    """배치 처리"""
    # 메타데이터 매퍼 설정
    metadata_mapper = None
    if args.metadata:
        mapper = MetadataMapper(args.metadata)
        metadata_mapper = mapper

    # 배치 프로세서 설정
    processor = BatchProcessor(
        workers=args.workers,
        timeout=args.timeout,
        metadata_mapper=metadata_mapper,
    )

    # 파일 목록 수집
    if args.filelist:
        with open(args.filelist, "r", encoding="utf-8") as f:
            files = [line.strip() for line in f if line.strip()]
    else:
        files = []
        for directory in args.directories:
            path = Path(directory)
            if args.recursive:
                files.extend(str(f) for f in path.rglob("*.hwp"))
            else:
                files.extend(str(f) for f in path.glob("*.hwp"))

    if not files:
        print("✗ 처리할 HWP 파일이 없음", file=sys.stderr)
        return 1

    print(f"📁 {len(files)}개 파일 처리 시작...")

    # 처리 실행
    result = processor.process_files(files, progress=not args.quiet)

    # 결과 출력
    print(f"\n📊 결과: 성공 {result.success}/{result.total} ({result.success_rate:.1%})")

    # 출력 저장
    if args.output:
        output_dir = Path(args.output)
        output_dir.mkdir(parents=True, exist_ok=True)

        exporter = YAMLExporter(str(output_dir))

        if args.format == "yaml":
            saved = exporter.export_batch(result, metadata_mapper)
            print(f"💾 YAML 저장: {len(saved)}개 → {output_dir}")
        else:  # jsonl
            output_file = output_dir / "training_data.jsonl"
            count = exporter.export_batch_jsonl(result, str(output_file), metadata_mapper)
            print(f"💾 JSONL 저장: {count}개 → {output_file}")

        # 실패 로그
        if result.failed > 0:
            failed_log = output_dir / "failed.jsonl"
            exporter.export_failed_log(result, str(failed_log))
            print(f"📝 실패 로그: {result.failed}개 → {failed_log}")

    return 0 if result.failed == 0 else 1


def cmd_info(args):
    """파일 정보 출력"""
    from .reader import HWPReader, HWPReaderError

    try:
        with HWPReader(args.file) as reader:
            meta = reader.metadata

            print(f"파일: {meta.filename}")
            print(f"버전: {meta.version}")
            print(f"압축: {'예' if meta.is_compressed else '아니오'}")
            print(f"크기: {meta.file_size_bytes:,} bytes")
            print(f"스트림: {len(meta.streams)}개")

            if args.verbose:
                print("\n스트림 목록:")
                for stream in meta.streams:
                    print(f"  - {stream}")

        return 0

    except HWPReaderError as e:
        print(f"✗ 오류: {e}", file=sys.stderr)
        return 1


def main():
    """CLI 진입점"""
    parser = argparse.ArgumentParser(
        prog="hwp-parser",
        description="HWP 5.x 텍스트 추출기 (olefile 기반)",
    )
    subparsers = parser.add_subparsers(dest="command", help="명령")

    # extract 명령
    p_extract = subparsers.add_parser("extract", help="단일 파일 텍스트 추출")
    p_extract.add_argument("file", help="HWP 파일 경로")
    p_extract.add_argument("-o", "--output", help="출력 파일 경로")
    p_extract.set_defaults(func=cmd_extract)

    # batch 명령
    p_batch = subparsers.add_parser("batch", help="배치 처리")
    p_batch.add_argument("directories", nargs="*", help="HWP 디렉토리")
    p_batch.add_argument("-f", "--filelist", help="파일 목록 텍스트")
    p_batch.add_argument("-o", "--output", help="출력 디렉토리")
    p_batch.add_argument("-m", "--metadata", help="메타데이터 JSONL 파일")
    p_batch.add_argument("-w", "--workers", type=int, default=None, help="워커 수")
    p_batch.add_argument("-t", "--timeout", type=int, default=30, help="타임아웃(초)")
    p_batch.add_argument("-r", "--recursive", action="store_true", help="재귀 탐색")
    p_batch.add_argument("-q", "--quiet", action="store_true", help="진행률 숨김")
    p_batch.add_argument("--format", choices=["yaml", "jsonl"], default="jsonl", help="출력 형식")
    p_batch.set_defaults(func=cmd_batch)

    # info 명령
    p_info = subparsers.add_parser("info", help="파일 정보 출력")
    p_info.add_argument("file", help="HWP 파일 경로")
    p_info.add_argument("-v", "--verbose", action="store_true", help="상세 출력")
    p_info.set_defaults(func=cmd_info)

    args = parser.parse_args()

    if args.command is None:
        parser.print_help()
        return 1

    return args.func(args)


if __name__ == "__main__":
    sys.exit(main())
