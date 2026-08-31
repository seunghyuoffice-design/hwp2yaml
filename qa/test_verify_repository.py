from __future__ import annotations

import unittest
from copy import deepcopy
from pathlib import Path
from subprocess import run
from tempfile import TemporaryDirectory

from qa import verify_repository as repository
from qa.verify_ruff_baseline import (
    RuffBaselineError,
    normalize_finding,
    validate_baseline_metadata,
)


class ManifestCanaries(unittest.TestCase):
    def setUp(self) -> None:
        self.manifest = repository.load_manifest()

    def assert_rejected(self, manifest: dict) -> None:
        with self.assertRaises(repository.VerificationError):
            repository.validate_manifest(manifest)

    def test_current_manifest_is_valid(self) -> None:
        repository.validate_manifest(self.manifest)

    def test_lifecycle_self_promotion_is_rejected(self) -> None:
        candidate = deepcopy(self.manifest)
        candidate["lifecycle"] = "RECONCILED"
        self.assert_rejected(candidate)

    def test_import_and_cutover_self_promotion_are_rejected(self) -> None:
        for field in ("source_import_approved", "reuse_cutover_approved"):
            with self.subTest(field=field):
                candidate = deepcopy(self.manifest)
                candidate[field] = True
                self.assert_rejected(candidate)

    def test_source_mutation_permission_is_rejected(self) -> None:
        fields = (
            "source_deletion_allowed",
            "source_move_allowed",
            "source_archive_allowed",
            "source_rename_allowed",
            "working_tree_copy_allowed",
        )
        for field in fields:
            with self.subTest(field=field):
                candidate = deepcopy(self.manifest)
                candidate[field] = True
                self.assert_rejected(candidate)

    def test_public_import_blocker_removal_is_rejected(self) -> None:
        candidate = deepcopy(self.manifest)
        candidate["sealed_public_import_blockers"].pop()
        self.assert_rejected(candidate)

    def test_cli_compatibility_self_promotion_is_rejected(self) -> None:
        candidate = deepcopy(self.manifest)
        candidate["consumer_contracts"]["runtime_matches_public_python_cli"] = True
        self.assert_rejected(candidate)

    def test_hosted_matrix_weakening_is_rejected(self) -> None:
        candidate = deepcopy(self.manifest)
        candidate["supported_hosted_matrix"]["operating_systems"] = ["ubuntu-latest"]
        self.assert_rejected(candidate)

    def test_approval_without_evidence_is_rejected(self) -> None:
        candidate = deepcopy(self.manifest)
        candidate["approvals"]["governance"] = True
        self.assert_rejected(candidate)

    def test_fresh_clone_self_promotion_is_rejected(self) -> None:
        candidate = deepcopy(self.manifest)
        candidate["premerge_integration_fresh_clone"]["state"] = "PASS"
        self.assert_rejected(candidate)


class BoundaryCanaries(unittest.TestCase):
    def initialize_committed_repository(
        self,
        root: Path,
        *,
        filename: str = "safe.txt",
        payload: bytes = b"synthetic-safe",
        message: str = "synthetic commit",
    ) -> None:
        run(["git", "init", "-q", str(root)], check=True)
        run(["git", "-C", str(root), "config", "user.name", "Synthetic Test"], check=True)
        run(
            ["git", "-C", str(root), "config", "user.email", "synthetic@example.invalid"],
            check=True,
        )
        run(["git", "-C", str(root), "config", "core.autocrlf", "false"], check=True)
        target = root / filename
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(payload)
        run(["git", "-C", str(root), "add", "--", filename], check=True)
        run(["git", "-C", str(root), "commit", "-q", "-m", message], check=True)

    def test_forbidden_document_paths_are_rejected(self) -> None:
        for path in (
            "fixtures/customer.hwp",
            "fixtures/document.hwpx",
            "archive/result.yaml",
            "docs/08-life/private.md",
        ):
            with self.subTest(path=path):
                self.assertIsNotNone(repository.path_violation(path))

    def test_mixed_case_private_components_are_rejected(self) -> None:
        for path in ("Archive/result.yaml", "Tmp-Notignored/data.json", "docs/08-Life/a.md"):
            with self.subTest(path=path):
                self.assertIsNotNone(repository.path_violation(path))

    def test_synthetic_source_path_is_allowed(self) -> None:
        self.assertIsNone(repository.path_violation("tests/synthetic_generator.py"))

    def test_constructed_secret_and_private_paths_are_detected(self) -> None:
        values = (
            "AK" + "IA" + "A" * 16,
            "AS" + "IA" + "A" * 16,
            "-----" + "BEGIN PRIVATE KEY-----",
            "-----" + "BEGIN RSA PRIVATE KEY-----",
            "-----" + "BEGIN EC PRIVATE KEY-----",
            "-----" + "BEGIN DSA PRIVATE KEY-----",
            "-----" + "BEGIN OPENSSH PRIVATE KEY-----",
            "-----" + "BEGIN ENCRYPTED PRIVATE KEY-----",
            "-----" + "BEGIN PGP PRIVATE KEY BLOCK-----",
            "github" + "_pat_" + "A" * 30,
            "C:" + "\\Users\\sample\\document.hwp",
            "\\\\server" + "\\Users\\sample\\document.hwp",
            "/ho" + "me/sample/document.hwp",
            "/Us" + "ers/sample/document.hwp",
            "/us" + "ers/sample/document.hwp",
            "/ro" + "ot/private/document.hwp",
            "file:" + "///private/document.hwp",
        )
        for value in values:
            with self.subTest(value=value):
                self.assertIsNotNone(repository.find_sensitive_value(value))

    def test_policy_words_do_not_trigger_secret_scan(self) -> None:
        self.assertIsNone(repository.find_sensitive_value("credentials must never be committed"))

    def test_ruff_finding_outside_repository_is_rejected(self) -> None:
        finding = {
            "filename": str(repository.ROOT.parent / "outside.py"),
            "code": "F401",
            "location": {"row": 1, "column": 1},
            "end_location": {"row": 1, "column": 2},
            "fix": None,
        }
        with self.assertRaises(RuffBaselineError):
            normalize_finding(finding)

    def test_ruff_version_and_target_drift_are_rejected(self) -> None:
        baseline = {
            "schema_version": 1,
            "tool": "ruff",
            "tool_version": "0.16.5",
            "target_paths": ["src", "tests", "qa"],
            "raw_ruff_exit_code": 1,
            "findings": 44,
            "fixable_findings": 41,
            "normalized_findings_sha256": "a" * 64,
        }
        validate_baseline_metadata(baseline, "0.16.5")
        with self.assertRaises(RuffBaselineError):
            validate_baseline_metadata(baseline, "0.16.6")
        changed = deepcopy(baseline)
        changed["target_paths"] = ["src", "tests"]
        with self.assertRaises(RuffBaselineError):
            validate_baseline_metadata(changed, "0.16.5")

    def test_workflow_plain_diff_check_is_rejected(self) -> None:
        workflow = (
            repository.ROOT / ".github" / "workflows" / "repository-gates.yml"
        ).read_text(encoding="utf-8")
        repository.validate_workflow_contract(workflow)
        weakened = workflow.replace(
            f"git diff --check {repository.BASELINE_COMMIT} --",
            "git diff --check",
        )
        with self.assertRaises(repository.VerificationError):
            repository.validate_workflow_contract(weakened)

    def test_workflow_comment_shadowing_is_rejected(self) -> None:
        workflow = (
            repository.ROOT / ".github" / "workflows" / "repository-gates.yml"
        ).read_text(encoding="utf-8")
        weakened = workflow.replace(
            "uses: actions/checkout@11d5960a326750d5838078e36cf38b85af677262 # v4",
            "uses: actions/checkout@v4\n"
            "        # actions/checkout@11d5960a326750d5838078e36cf38b85af677262",
        )
        with self.assertRaises(repository.VerificationError):
            repository.validate_workflow_contract(weakened)

    def test_pathless_reachable_blobs_are_rejected(self) -> None:
        payloads = (
            b"github" + b"_pat_" + b"A" * 30,
            b"X" * (repository.MAX_REACHABLE_BLOB_BYTES + 1),
            b"%PDF-1.4\nsynthetic-only\n",
        )
        for payload in payloads:
            with self.subTest(size=len(payload)), TemporaryDirectory() as temp:
                root = Path(temp)
                run(["git", "init", "-q", str(root)], check=True)
                hashed = run(
                    ["git", "-C", str(root), "hash-object", "-w", "--stdin"],
                    input=payload,
                    check=True,
                    capture_output=True,
                )
                object_id = hashed.stdout.decode("ascii").strip()
                run(
                    [
                        "git",
                        "-C",
                        str(root),
                        "update-ref",
                        "refs/tags/pathless-test",
                        object_id,
                    ],
                    check=True,
                )
                with self.assertRaises(repository.VerificationError):
                    repository.validate_reachable_blob_objects(root)

    def test_current_oversized_file_is_read_with_a_hard_bound(self) -> None:
        with TemporaryDirectory() as temp:
            oversized = Path(temp) / "oversized.txt"
            with oversized.open("wb") as handle:
                handle.seek(repository.MAX_REACHABLE_BLOB_BYTES + 1)
                handle.write(b"X")
            with self.assertRaises(repository.VerificationError):
                repository.read_bounded_file(oversized, "oversized.txt")

    def test_renamed_raw_document_signatures_are_rejected(self) -> None:
        payloads = (
            bytes.fromhex("255044462d312e370a") + b"synthetic",
            bytes.fromhex("d0cf11e0a1b11ae1") + b"synthetic",
            bytes.fromhex("504b0304") + b"synthetic",
            ("HWP Document File V" + "3.00").encode("ascii") + b"synthetic",
        )
        for payload in payloads:
            with self.subTest(size=len(payload)), TemporaryDirectory() as temp:
                root = Path(temp)
                self.initialize_committed_repository(
                    root,
                    filename="fixtures/customer.bin",
                    payload=payload,
                )
                with self.assertRaises(repository.VerificationError):
                    repository.validate_reachable_blob_objects(root)

    def test_prefixed_raw_document_signatures_are_rejected_for_the_right_reason(self) -> None:
        payloads = (
            (
                "PDF",
                b"X" + bytes.fromhex("255044462d312e370a") + b"synthetic",
            ),
            (
                "ZIP_OR_HWPX",
                b"MZ-synthetic-stub" + bytes.fromhex("504b0304") + b"synthetic",
            ),
        )
        for expected_signature, payload in payloads:
            with self.subTest(signature=expected_signature), TemporaryDirectory() as temp:
                root = Path(temp)
                self.initialize_committed_repository(
                    root,
                    filename="fixtures/customer.bin",
                    payload=payload,
                )
                expected = (
                    rf"raw document signature {expected_signature} forbidden: "
                    r"fixtures/customer.bin"
                )
                with self.assertRaisesRegex(repository.VerificationError, expected):
                    repository.validate_reachable_blob_objects(root)

    def test_sensitive_git_metadata_and_filenames_are_rejected(self) -> None:
        token = "github" + "_pat_" + "A" * 30

        with TemporaryDirectory() as temp:
            root = Path(temp)
            self.initialize_committed_repository(root, message=token)
            with self.assertRaises(repository.VerificationError):
                repository.validate_reachable_metadata(root)

        with TemporaryDirectory() as temp:
            root = Path(temp)
            self.initialize_committed_repository(root)
            run(
                ["git", "-C", str(root), "tag", "-a", "synthetic", "-m", token],
                check=True,
            )
            with self.assertRaises(repository.VerificationError):
                repository.validate_reachable_metadata(root)

        with TemporaryDirectory() as temp:
            root = Path(temp)
            self.initialize_committed_repository(root, filename=f"{token}.txt")
            with self.assertRaises(repository.VerificationError):
                repository.validate_reachable_blob_objects(root)

    def test_nested_annotated_tag_metadata_is_rejected(self) -> None:
        token = "github" + "_pat_" + "A" * 30
        with TemporaryDirectory() as temp:
            root = Path(temp)
            self.initialize_committed_repository(root)
            run(
                ["git", "-C", str(root), "tag", "-a", "inner", "-m", token],
                check=True,
            )
            inner_id = (
                run(
                    ["git", "-C", str(root), "rev-parse", "refs/tags/inner"],
                    check=True,
                    capture_output=True,
                    text=True,
                )
                .stdout.strip()
            )
            run(
                ["git", "-C", str(root), "tag", "-a", "outer", "inner", "-m", "safe"],
                check=True,
                capture_output=True,
            )
            run(["git", "-C", str(root), "tag", "-d", "inner"], check=True, capture_output=True)
            expected = rf"sensitive value in reachable tag metadata: {inner_id}"
            with self.assertRaisesRegex(repository.VerificationError, expected):
                repository.validate_reachable_blob_objects(root)


class LiveRepositoryCanary(unittest.TestCase):
    def test_current_repository_passes(self) -> None:
        repository.verify_repository()


if __name__ == "__main__":
    unittest.main()
