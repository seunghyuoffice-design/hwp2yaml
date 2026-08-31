from __future__ import annotations

import hashlib
import json
import subprocess
import sys
from importlib import metadata
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
BASELINE_PATH = ROOT / "qa" / "ruff-baseline.v1.json"


class RuffBaselineError(RuntimeError):
    """Raised when the exact, acknowledged Ruff debt changes."""


def normalize_finding(item: dict[str, Any]) -> dict[str, Any]:
    filename = Path(item["filename"]).resolve()
    try:
        relative = filename.relative_to(ROOT).as_posix()
    except ValueError as exc:
        raise RuffBaselineError(f"Ruff finding escaped repository: {filename}") from exc
    location = item["location"]
    end_location = item["end_location"]
    return {
        "path": relative,
        "code": item["code"],
        "row": location["row"],
        "column": location["column"],
        "end_row": end_location["row"],
        "end_column": end_location["column"],
        "fixable": item.get("fix") is not None,
    }


def normalized_digest(items: list[dict[str, Any]]) -> tuple[str, list[dict[str, Any]]]:
    normalized = sorted(
        (normalize_finding(item) for item in items),
        key=lambda item: (
            item["path"],
            item["row"],
            item["column"],
            item["code"],
        ),
    )
    payload = json.dumps(
        normalized,
        ensure_ascii=True,
        separators=(",", ":"),
        sort_keys=True,
    ).encode("utf-8")
    return hashlib.sha256(payload).hexdigest(), normalized


def load_baseline() -> dict[str, Any]:
    try:
        data = json.loads(BASELINE_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise RuffBaselineError(f"cannot parse Ruff baseline: {exc}") from exc
    if not isinstance(data, dict):
        raise RuffBaselineError("Ruff baseline must be an object")
    return data


def validate_baseline_metadata(data: dict[str, Any], observed_version: str) -> None:
    expected = {
        "schema_version": 1,
        "tool": "ruff",
        "tool_version": "0.16.5",
        "target_paths": ["src", "tests", "qa"],
        "raw_ruff_exit_code": 1,
        "findings": 44,
        "fixable_findings": 41,
    }
    for field, value in expected.items():
        if data.get(field) != value:
            raise RuffBaselineError(f"Ruff baseline metadata drift: {field}")
    if observed_version != data["tool_version"]:
        raise RuffBaselineError(
            f"Ruff version drift: expected {data['tool_version']}, got {observed_version}"
        )
    digest = data.get("normalized_findings_sha256")
    if not isinstance(digest, str) or not re_full_sha256(digest):
        raise RuffBaselineError("invalid Ruff baseline fingerprint")


def re_full_sha256(value: str) -> bool:
    return len(value) == 64 and all(character in "0123456789abcdef" for character in value)


def run_ruff() -> tuple[subprocess.CompletedProcess[str], list[dict[str, Any]]]:
    completed = subprocess.run(
        [
            sys.executable,
            "-m",
            "ruff",
            "check",
            "--output-format",
            "json",
            "src",
            "tests",
            "qa",
        ],
        cwd=ROOT,
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if completed.returncode not in {0, 1}:
        raise RuffBaselineError(completed.stderr.strip() or "Ruff execution failed")
    try:
        items = json.loads(completed.stdout)
    except json.JSONDecodeError as exc:
        raise RuffBaselineError(f"invalid Ruff JSON: {exc}") from exc
    if not isinstance(items, list):
        raise RuffBaselineError("Ruff output must be a list")
    return completed, items


def verify_ruff_baseline() -> None:
    baseline = load_baseline()
    try:
        observed_version = metadata.version("ruff")
    except metadata.PackageNotFoundError as exc:
        raise RuffBaselineError("Ruff distribution is not installed") from exc
    validate_baseline_metadata(baseline, observed_version)
    completed, items = run_ruff()
    digest, normalized = normalized_digest(items)
    fixable = sum(1 for item in normalized if item["fixable"])
    expected_count = baseline.get("findings")
    expected_fixable = baseline.get("fixable_findings")
    expected_digest = baseline.get("normalized_findings_sha256")
    if completed.returncode != baseline.get("raw_ruff_exit_code"):
        raise RuffBaselineError("raw Ruff exit-code drift")
    if len(normalized) != expected_count:
        raise RuffBaselineError(f"Ruff finding count drift: {len(normalized)}")
    if fixable != expected_fixable:
        raise RuffBaselineError(f"Ruff fixable count drift: {fixable}")
    if digest != expected_digest:
        raise RuffBaselineError(f"Ruff finding fingerprint drift: {digest}")
    if expected_count != 44 or completed.returncode != 1:
        raise RuffBaselineError("baseline must preserve the acknowledged 44-finding HOLD")


def main() -> int:
    try:
        verify_ruff_baseline()
    except RuffBaselineError as exc:
        print(f"[FAIL] {exc}")
        return 1
    print("[PASS] Ruff regression floor: 44 acknowledged findings, release gate still failing")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
