from __future__ import annotations

import hashlib
import json
import re
import subprocess
from copy import deepcopy
from pathlib import Path, PurePosixPath
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
BASELINE_COMMIT = "56231e2e4e26e3aff3dcff6ef82f5789aaa4a9d8"
GOVERNANCE_COMMIT = "6f6ebd8ce9cfe27e6331a3afa3e38971fe496e45"
EXPECTED_WORKFLOW_BLOB = "36d902da98d446525e4b69eeecd019a7fc0d0e73"

GOVERNANCE_BLOBS = {
    "AGENTS.md": "534c6d7e90682d022ca0b5b337df4e970925809d",
    "SECURITY.md": "4da048490b781bfe08f89c2600d5318b5398df2c",
    "docs/OWNERSHIP.md": "a2c10d0fd14631c643ff3abb56be4f3583aa361e",
    "ownership/reuse-reconciliation.v1.json": (
        "4ae575242856d51615921d283cd4945c2d9986cf"
    ),
}

VERIFY_PATHS = {
    ".github/workflows/repository-gates.yml",
    "docs/VERIFICATION.md",
    "qa/__init__.py",
    "qa/ruff-baseline.v1.json",
    "qa/test_verify_repository.py",
    "qa/verify_repository.py",
    "qa/verify_ruff_baseline.py",
}
EXPECTED_ADDITIONS = set(GOVERNANCE_BLOBS) | VERIFY_PATHS

EXPECTED_CONSUMER_BLOBS = {
    "config/module-boundaries.toml": "12eda31c31952cd9b7ccdfb474697ef3dc6c612b",
    "config/ingest_config.yaml": "536cd0c60631ca93286f3f512692c69459a7edaf",
    "config/pipeline_runtime.yaml": "7458290b73d3acc95a64696b74f205fef8837b36",
    "tools/rust-workers/crates/conversion-port/Cargo.toml": (
        "33cfb12b67db84f0f5daad7b2ccd2bd15654a9b7"
    ),
    "tools/rust-workers/crates/conversion-port/src/lib.rs": (
        "4ff264145cb6fb6d4e217d96864f9b39e380989c"
    ),
    "tools/rust-workers/crates/pipeline-core/Cargo.toml": (
        "17905bc4cf9eafe2bbf9ec7ec4460d8149af0e83"
    ),
    "tools/rust-workers/crates/pipeline-core/src/parsers/unified_parser.rs": (
        "68d8fd057b2f0cb5d49b6728b0b5ba2fa73db08f"
    ),
    "tools/pipeline-rs/crates/pipeline-runtime/src/hwp_stage.rs": (
        "0278d8a9f8f52a025d6f9421ec82a4590e6a3957"
    ),
}

EXPECTED_BLOCKERS = {
    "UNRESOLVED_THIRD_PARTY_COPYRIGHT_LICENSE_NOTICE_AND_GENERATION_CHAIN",
    "UNRESOLVED_OR_UNAPPROVED_DOCUMENT_FIXTURE_PROVENANCE",
    "PIPELINE_RELATIVE_CONVERSION_PORT_DEPENDENCY",
    "CLI_AND_SCHEMA_AUTHORITY_CONFLICT",
    "PARSER_RESOURCE_LIMITS_AND_CANCELLATION_GAPS",
    "EXTERNAL_PROCESS_MODEL_AND_NO_EGRESS_BOUNDARY_UNRESOLVED",
    "DEPENDENCY_ADVISORY_AND_SBOM_REVIEW_REQUIRED",
    "DETERMINISTIC_OUTPUT_AND_PATH_PRIVACY_UNPROVED",
}

REQUIRED_MARKERS = {
    "REPOSITORY_LOCAL_SUPREMACY",
    "NESTED_RULES_MAY_ONLY_STRENGTHEN",
    "PUBLIC_PYTHON_BASELINE_ONLY",
    "REUSE_IS_NOT_CUTOVER",
    "NO_DUPLICATE_REMOTE",
    "PUBLIC_DATA_FAIL_CLOSED",
    "IMPORT_PROVENANCE_REQUIRED",
    "NO_SOURCE_DELETION_WITHOUT_APPROVAL",
    "UNTRUSTED_INPUT_LIMITS_REQUIRED",
    "ARCHIVE_XML_FAIL_CLOSED",
    "EXTERNAL_PROCESS_OPT_IN",
    "DETERMINISTIC_OUTPUT_REQUIRED",
    "SINGLE_VERSION_AUTHORITY",
    "CLAIMS_REQUIRE_EVIDENCE",
    "HARNESS_EVOLUTION_REQUIRED",
    "AGENTS_HARNESS_COHERENCE",
    "POSITIVE_AND_NEGATIVE_BEHAVIOR_REQUIRED",
    "ROLE_SEPARATED_WORKTREES",
    "SUPPORTED_HOSTED_MATRIX",
    "EXTERNAL_CI_AND_FRESH_CLONE_REQUIRED",
}

FORBIDDEN_SUFFIXES = (
    ".hwp",
    ".hwp3",
    ".hwpx",
    ".pdf",
    ".xml.zst",
    ".ocr.txt",
)
FORBIDDEN_COMPONENTS = {"archive", "tmp-notignored", "08-life"}
MAX_REACHABLE_BLOB_BYTES = 1_000_000
RAW_DOCUMENT_ALLOWLIST: set[str] = set()


class VerificationError(RuntimeError):
    """Raised when a repository boundary or sealed receipt drifts."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise VerificationError(message)


def run_git_at(
    repository: Path, *args: str, check: bool = True
) -> subprocess.CompletedProcess[str]:
    completed = subprocess.run(
        ["git", "-C", str(repository), *args],
        check=False,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
    )
    if check and completed.returncode != 0:
        detail = completed.stderr.strip() or completed.stdout.strip()
        raise VerificationError(f"git {' '.join(args)} failed: {detail}")
    return completed


def run_git(*args: str, check: bool = True) -> subprocess.CompletedProcess[str]:
    return run_git_at(ROOT, *args, check=check)


def run_git_bytes_at(repository: Path, *args: str) -> subprocess.CompletedProcess[bytes]:
    completed = subprocess.run(
        ["git", "-C", str(repository), *args],
        check=False,
        capture_output=True,
    )
    if completed.returncode != 0:
        detail = (completed.stderr or completed.stdout).decode("utf-8", errors="replace").strip()
        raise VerificationError(f"git {' '.join(args)} failed: {detail}")
    return completed


def git_output(*args: str) -> str:
    return run_git(*args).stdout.strip()


def git_output_at(repository: Path, *args: str) -> str:
    return run_git_at(repository, *args).stdout.strip()


def nested_value(data: dict[str, Any], path: str) -> Any:
    value: Any = data
    for part in path.split("."):
        require(isinstance(value, dict) and part in value, f"missing manifest field: {path}")
        value = value[part]
    return value


def require_value(data: dict[str, Any], path: str, expected: Any) -> None:
    observed = nested_value(data, path)
    require(observed == expected, f"manifest drift at {path}: {observed!r}")


def load_manifest() -> dict[str, Any]:
    path = ROOT / "ownership" / "reuse-reconciliation.v1.json"
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise VerificationError(f"cannot parse reconciliation manifest: {exc}") from exc
    require(isinstance(data, dict), "reconciliation manifest must be an object")
    return data


def validate_manifest(data: dict[str, Any]) -> None:
    require_value(data, "schema_version", 1)
    require_value(data, "receipt_kind", "HWP_REUSE_RECONCILIATION")
    require_value(data, "lifecycle", "HOLD_DEPENDENCY_RECONCILE")
    require_value(data, "selected_reuse_target", "seunghyuoffice-design/hwp2yaml")
    require_value(
        data,
        "authority_order",
        [
            "AGENTS.md",
            "ownership/reuse-reconciliation.v1.json",
            "docs/OWNERSHIP.md",
            "docs/VERIFICATION.md",
        ],
    )

    for field in (
        "new_duplicate_remote_allowed",
        "reuse_cutover_approved",
        "source_import_approved",
        "source_deletion_allowed",
        "source_move_allowed",
        "source_archive_allowed",
        "source_rename_allowed",
        "working_tree_copy_allowed",
    ):
        require_value(data, field, False)
    require_value(data, "role_worktrees_required", True)

    require_value(data, "public_target.commit", BASELINE_COMMIT)
    require_value(
        data,
        "public_target.root_tree",
        "0663b1cdb86d2b73dae1c25433a0f9ddefd06c1d",
    )
    require_value(
        data,
        "public_target.source_tree",
        "609e7ef07890e4174d4f761e2efed8f26bc58481",
    )
    require_value(
        data,
        "public_target.tests_tree",
        "e339e865becd74829437a62ee1bb8e58f6be6a26",
    )
    require_value(data, "public_target.tracked_files", 21)
    require_value(data, "public_target.declared_package_version", "0.6.1")
    require_value(data, "public_target.importable_version", "0.6.0")
    require_value(data, "public_target.release_authority", "CURRENT_BASELINE_ONLY")

    candidates = nested_value(data, "candidate_implementations")
    require(isinstance(candidates, list) and len(candidates) == 2, "candidate set drift")
    by_repository = {candidate.get("repository"): candidate for candidate in candidates}
    pipeline = by_repository.get("seunghyuoffice-design/dyarchy-pipeline")
    dya = by_repository.get("seunghyuoffice-design/Dyarchy-v6")
    require(isinstance(pipeline, dict), "pipeline candidate missing")
    require(isinstance(dya, dict), "Dya legacy evidence missing")
    require(
        pipeline.get("commit") == "b565aa04737389e21a4113a2ec2af0ca1fb68c76",
        "pipeline commit drift",
    )
    require(pipeline.get("component_files") == 33, "pipeline component count drift")
    require(pipeline.get("public_import_allowed") is False, "pipeline import self-promotion")
    require(pipeline.get("standalone_build_ready") is False, "standalone build self-promotion")
    require(dya.get("commit") == "52f80a13631e3054b76c4f9fd0df7a8c28161176", "Dya commit drift")
    require(dya.get("component_files") == 32, "Dya component count drift")
    require(dya.get("authority") == "LEGACY_EVIDENCE_ONLY", "Dya authority drift")
    require(dya.get("working_tree_read_allowed") is False, "Dya working-tree access enabled")
    require(dya.get("public_import_allowed") is False, "Dya import self-promotion")

    require_value(
        data,
        "consumer_contracts.exact_path_blobs",
        EXPECTED_CONSUMER_BLOBS,
    )
    require_value(data, "consumer_contracts.runtime_matches_public_python_cli", False)
    require_value(data, "consumer_contracts.runtime_matches_pipeline_rust_main_cli", False)
    require(
        set(nested_value(data, "sealed_public_import_blockers")) == EXPECTED_BLOCKERS,
        "sealed public-import blocker set drift",
    )
    require_value(
        data,
        "excluded_candidate_paths",
        ["src/hwp2yaml-rs/test_files/1999-046.yaml"],
    )
    require_value(data, "fixture_policy.synthetic_only_by_default", True)
    require_value(data, "fixture_policy.generator_and_hash_required", True)
    require_value(data, "fixture_policy.raw_private_documents_allowed", False)
    require_value(data, "fixture_policy.absolute_or_file_uri_source_paths_allowed", False)

    require_value(
        data,
        "supported_hosted_matrix.operating_systems",
        ["ubuntu-latest", "windows-latest"],
    )
    require_value(data, "supported_hosted_matrix.python_versions", ["3.10", "3.12"])
    require_value(data, "supported_hosted_matrix.rust_artifact_supported", False)
    require_value(data, "machine_validation.required", True)
    require_value(data, "machine_validation.verifier_path", "qa/verify_repository.py")
    require_value(data, "machine_validation.state", "PENDING_VERIFY_ROLE_BRANCH")
    require_value(data, "hosted_ci.required", True)
    require_value(data, "hosted_ci.required_jobs_must_be_created", True)
    require_value(data, "hosted_ci.integration_head", None)
    require_value(data, "hosted_ci.run_ids", [])
    require_value(data, "hosted_ci.state", "UNVERIFIED_NO_WORKFLOW_AT_PUBLIC_BASELINE")
    require_value(
        data,
        "approvals",
        {
            "governance": False,
            "verification": False,
            "security": False,
            "release": False,
            "evidence": [],
        },
    )
    require_value(data, "premerge_integration_fresh_clone.required", True)
    require_value(data, "premerge_integration_fresh_clone.hosted_ci_head_must_match", True)
    require_value(data, "premerge_integration_fresh_clone.integration_head", None)
    require_value(data, "premerge_integration_fresh_clone.evidence", None)
    require_value(data, "premerge_integration_fresh_clone.state", "UNVERIFIED")
    require_value(
        data,
        "postmerge_final_main_fresh_clone.required_before_reconciled_release_or_cutover",
        True,
    )
    require_value(data, "postmerge_final_main_fresh_clone.main_commit", None)
    require_value(data, "postmerge_final_main_fresh_clone.evidence", None)
    require_value(data, "postmerge_final_main_fresh_clone.state", "UNVERIFIED")
    require_value(
        data,
        "repository_security_setting.private_vulnerability_reporting",
        "ENABLED_VERIFIED_2026-08-31",
    )


def validate_constitution() -> None:
    text = (ROOT / "AGENTS.md").read_text(encoding="utf-8")
    missing = sorted(marker for marker in REQUIRED_MARKERS if marker not in text)
    require(not missing, f"constitution marker missing: {missing}")
    require(
        "하위 `AGENTS.md`" in text and "강화" in text,
        "nested AGENTS strengthening rule missing",
    )


def validate_governance_blobs() -> None:
    for path, expected in GOVERNANCE_BLOBS.items():
        observed = git_output("hash-object", f"--path={path}", "--", path)
        require(observed == expected, f"sealed governance blob drift: {path}")


def parse_name_status(text: str) -> dict[str, str]:
    changes: dict[str, str] = {}
    for raw_line in text.splitlines():
        if not raw_line:
            continue
        fields = raw_line.split("\t")
        require(len(fields) == 2, f"rename/copy or malformed diff status: {raw_line}")
        status, path = fields
        require(status == "A", f"only additive governance/verify paths allowed: {raw_line}")
        changes[path] = status
    return changes


def validate_repository_diff() -> None:
    require(git_output("cat-file", "-t", BASELINE_COMMIT) == "commit", "baseline missing")
    require(git_output("cat-file", "-t", GOVERNANCE_COMMIT) == "commit", "governance missing")
    require(
        run_git("merge-base", "--is-ancestor", BASELINE_COMMIT, "HEAD", check=False).returncode
        == 0,
        "baseline is not an ancestor of HEAD",
    )
    require(
        run_git("merge-base", "--is-ancestor", GOVERNANCE_COMMIT, "HEAD", check=False).returncode
        == 0,
        "governance commit is not an ancestor of HEAD",
    )
    changes = parse_name_status(git_output("diff", "--name-status", BASELINE_COMMIT, "--"))
    require(set(changes) == EXPECTED_ADDITIONS, f"unexpected path diff: {sorted(changes)}")

    tracked = set(git_output("ls-files").splitlines())
    require(EXPECTED_ADDITIONS <= tracked, "required governance/verify path not tracked")

    status_lines = run_git("status", "--porcelain=v1", "--untracked-files=all").stdout.splitlines()
    for line in status_lines:
        require(not line.startswith("??"), f"untracked path forbidden: {line[3:]}")
        require(len(line) >= 4 and line[1] == " ", f"unstaged change forbidden: {line}")
        require(line[3:] in VERIFY_PATHS, f"unexpected staged path: {line}")


def path_violation(path: str) -> str | None:
    lowered = path.lower()
    if lowered.endswith(FORBIDDEN_SUFFIXES):
        return "forbidden document/corpus suffix"
    parts = {part.lower() for part in PurePosixPath(path).parts}
    if parts & FORBIDDEN_COMPONENTS:
        return "forbidden private/archive path component"
    return None


def secret_patterns() -> tuple[re.Pattern[str], ...]:
    private_header = "-----" + "BEGIN "
    file_uri = "file:" + "//"
    root_home = "/ro" + "ot/"
    fine_grained_pat = "github" + "_pat_"
    pgp_private_header = "-----" + "BEGIN PGP PRIVATE KEY BLOCK-----"
    return (
        re.compile(r"AK" + r"IA[0-9A-Z]{16}"),
        re.compile(r"AS" + r"IA[0-9A-Z]{16}"),
        re.compile(r"gh[pousr]_[A-Za-z0-9]{20,}"),
        re.compile(re.escape(fine_grained_pat) + r"[A-Za-z0-9_]{20,}"),
        re.compile(
            re.escape(private_header)
            + r"(?:ENCRYPTED |RSA |EC |DSA |OPENSSH )?PRIVATE KEY-----"
        ),
        re.compile(re.escape(pgp_private_header)),
        re.compile(r"[A-Za-z]:[\\/]+Users[\\/]+[A-Za-z0-9._-]+", re.IGNORECASE),
        re.compile(r"\\\\[^\\\r\n]+\\Users\\[A-Za-z0-9._-]+\\", re.IGNORECASE),
        re.compile(r"/home/[A-Za-z0-9._-]+/"),
        re.compile(r"/Users/[A-Za-z0-9._-]+/", re.IGNORECASE),
        re.compile(re.escape(root_home)),
        re.compile(re.escape(file_uri), re.IGNORECASE),
    )


def find_sensitive_value(text: str) -> str | None:
    for pattern in secret_patterns():
        if pattern.search(text):
            return pattern.pattern
    return None


def raw_document_signature(payload: bytes) -> str | None:
    if bytes.fromhex("255044462d") in payload[:1024]:
        return "PDF"
    if bytes.fromhex("d0cf11e0a1b11ae1") in payload:
        return "OLE_OR_HWP5"
    zip_signatures = (
        bytes.fromhex("504b0304"),
        bytes.fromhex("504b0506"),
        bytes.fromhex("504b0708"),
    )
    if any(signature in payload for signature in zip_signatures):
        return "ZIP_OR_HWPX"
    if ("HWP Document File V" + "3.00").encode("ascii") in payload:
        return "HWP3"
    return None


def reject_raw_document(payload: bytes, path: str) -> None:
    signature = raw_document_signature(payload)
    require(
        signature is None or path in RAW_DOCUMENT_ALLOWLIST,
        f"raw document signature {signature} forbidden: {path}",
    )


def git_blob_id(text: str) -> str:
    payload = text.encode("utf-8")
    header = f"blob {len(payload)}\0".encode("ascii")
    return hashlib.sha1(header + payload).hexdigest()


def validate_workflow_contract(text: str) -> None:
    require(
        git_blob_id(text) == EXPECTED_WORKFLOW_BLOB,
        "workflow byte contract drift",
    )
    required_fragments = {
        "permissions:\n  contents: read",
        "os: [ubuntu-latest, windows-latest]",
        'python: ["3.10", "3.12"]',
        "actions/checkout@11d5960a326750d5838078e36cf38b85af677262",
        "actions/setup-python@a26af69be951a213d495a4c3e4e4022e16d87065",
        f"git diff --check {BASELINE_COMMIT} --",
    }
    missing = sorted(fragment for fragment in required_fragments if fragment not in text)
    require(not missing, f"workflow contract drift: {missing}")


def reachable_commits() -> list[str]:
    commits = reachable_commits_at(ROOT)
    require(bool(commits), "no reachable commits")
    return commits


def read_bounded_file(path: Path, label: str) -> bytes:
    with path.open("rb") as handle:
        payload = handle.read(MAX_REACHABLE_BLOB_BYTES + 1)
    require(
        len(payload) <= MAX_REACHABLE_BLOB_BYTES,
        f"oversized current file: {label}",
    )
    return payload


def reachable_commits_at(repository: Path) -> list[str]:
    candidates = git_output_at(repository, "rev-list", "--all").splitlines()
    commits = {
        object_id
        for object_id in candidates
        if git_output_at(repository, "cat-file", "-t", object_id) == "commit"
    }
    return sorted(commits)


def tree_blob_paths_at(repository: Path) -> dict[str, set[str]]:
    paths_by_blob: dict[str, set[str]] = {}
    for commit in reachable_commits_at(repository):
        tree = run_git_at(repository, "ls-tree", "-r", commit).stdout.splitlines()
        for line in tree:
            metadata, path = line.split("\t", 1)
            _mode, object_type, object_id = metadata.split(" ")
            if object_type != "blob":
                continue
            paths_by_blob.setdefault(object_id, set()).add(path)
    return paths_by_blob


def validate_current_tracked_files() -> None:
    for line in git_output("ls-files", "--stage").splitlines():
        metadata, path = line.split("\t", 1)
        mode, object_id, stage = metadata.split(" ")
        require(mode == "100644", f"symlink, executable or gitlink forbidden: {path}")
        require(stage == "0", f"unmerged index entry forbidden: {path}")
        require(len(object_id) == 40, f"unexpected tracked object id: {path}")
        index_size = int(git_output("cat-file", "-s", object_id))
        require(
            index_size <= MAX_REACHABLE_BLOB_BYTES,
            f"oversized staged blob: {path}",
        )
        violation = path_violation(path)
        require(violation is None, f"{violation} in current tree: {path}")
        path_finding = find_sensitive_value(path)
        require(path_finding is None, f"sensitive value in current path: {path}")
        full_path = ROOT / path
        require(full_path.is_file(), f"tracked file missing from worktree: {path}")
        payload = read_bounded_file(full_path, path)
        finding = find_sensitive_value(payload.decode("utf-8", errors="replace"))
        require(finding is None, f"sensitive value in current file: {path}")
        reject_raw_document(payload, path)


def validate_reachable_history() -> None:
    for commit in reachable_commits():
        paths = git_output("ls-tree", "-r", "--name-only", commit).splitlines()
        for path in paths:
            violation = path_violation(path)
            require(violation is None, f"{violation} in reachable history: {path}")
            path_finding = find_sensitive_value(path)
            require(path_finding is None, f"sensitive value in reachable path: {path}")

    validate_reachable_metadata(ROOT)
    validate_reachable_blob_objects(ROOT)


def validate_reachable_metadata(repository: Path) -> None:
    for commit in reachable_commits_at(repository):
        raw_commit = run_git_at(repository, "cat-file", "commit", commit).stdout
        finding = find_sensitive_value(raw_commit)
        require(finding is None, f"sensitive value in reachable commit metadata: {commit}")

    tags = run_git_at(
        repository,
        "for-each-ref",
        "--format=%(refname)\t%(objectname)\t%(objecttype)",
        "refs/tags",
    ).stdout.splitlines()
    for line in tags:
        refname, object_id, object_type = line.split("\t")
        finding = find_sensitive_value(refname)
        require(finding is None, f"sensitive value in tag name: {refname}")
        if object_type == "tag":
            raw_tag = run_git_at(repository, "cat-file", "tag", object_id).stdout
            finding = find_sensitive_value(raw_tag)
            require(finding is None, f"sensitive value in annotated tag metadata: {refname}")


def validate_reachable_blob_objects(repository: Path) -> None:
    objects = run_git_at(repository, "rev-list", "--objects", "--all").stdout.splitlines()
    tree_paths = tree_blob_paths_at(repository)
    seen: set[str] = set()
    for line in objects:
        object_id = line.split(" ", 1)[0]
        if object_id in seen:
            continue
        seen.add(object_id)
        object_type = git_output_at(repository, "cat-file", "-t", object_id)
        if object_type == "tag":
            raw_tag = run_git_at(repository, "cat-file", "tag", object_id).stdout
            finding = find_sensitive_value(raw_tag)
            require(
                finding is None,
                f"sensitive value in reachable tag metadata: {object_id}",
            )
            continue
        if object_type != "blob":
            continue
        size = int(git_output_at(repository, "cat-file", "-s", object_id))
        actual_paths = tree_paths.get(object_id)
        require(actual_paths is not None, f"non-tree reachable blob forbidden: {object_id}")
        label = min(actual_paths)
        require(size <= MAX_REACHABLE_BLOB_BYTES, f"oversized reachable blob: {label}")
        blob = run_git_bytes_at(repository, "cat-file", "blob", object_id).stdout
        finding = find_sensitive_value(blob.decode("utf-8", errors="replace"))
        require(finding is None, f"sensitive value in reachable blob: {label}")
        for actual_path in actual_paths:
            violation = path_violation(actual_path)
            require(violation is None, f"{violation} in reachable blob: {actual_path}")
            path_finding = find_sensitive_value(actual_path)
            require(path_finding is None, f"sensitive value in reachable path: {actual_path}")
            reject_raw_document(blob, actual_path)


def verify_repository() -> None:
    validate_manifest(load_manifest())
    validate_constitution()
    validate_governance_blobs()
    validate_repository_diff()
    workflow = (ROOT / ".github" / "workflows" / "repository-gates.yml").read_text(
        encoding="utf-8"
    )
    validate_workflow_contract(workflow)
    validate_current_tracked_files()
    validate_reachable_history()


def mutated_manifest() -> dict[str, Any]:
    """Return an independent manifest copy for negative canaries."""
    return deepcopy(load_manifest())


def main() -> int:
    try:
        verify_repository()
    except VerificationError as exc:
        print(f"[FAIL] {exc}")
        return 1
    print(
        "[PASS] hwp2yaml repository boundary: "
        f"baseline={BASELINE_COMMIT[:12]} additions={len(EXPECTED_ADDITIONS)}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
