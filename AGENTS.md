# hwp2yaml 저장소 헌법

CONSTITUTION_VERSION: `1.0.0`

이 파일은 `hwp2yaml` 저장소 안에서 가장 높은 규칙입니다. 이 저장소는 공개 HWP/HWPX 변환 제품의 명시적으로 검증된 경계만 소유하며, Dyarchy-v6·ROVER·`dyarchy-pipeline` 또는 개인 문서의 집합소가 아닙니다.

## 우선순위

`REPOSITORY_LOCAL_SUPREMACY`

- 권위 순서는 이 최상위 `AGENTS.md`, `ownership/reuse-reconciliation.v1.json`, ownership/verification 문서, 그 밖의 파일 순입니다.
- JSON과 문서는 이 헌법을 기계화하고 설명할 뿐, 이 헌법의 금지·검증 조건을 완화하거나 승인 상태로 스스로 승격할 수 없습니다.
- 외부 monorepo의 편의, 같은 패키지 이름 또는 과거 운영 경로를 이유로 provenance, privacy, untrusted-input security, reproducibility 또는 release gate를 약화할 수 없습니다.

`NESTED_RULES_MAY_ONLY_STRENGTHEN`

- 하위 `AGENTS.md`는 해당 하위 영역의 규칙을 구체화하거나 강화할 수 있지만 이 헌법을 완화·삭제·우회할 수 없습니다.
- 충돌 시 더 강한 fail-closed, privacy, provenance, deterministic-output 및 evidence 조건을 적용합니다.

## 소유 경계와 단일 권위

`PUBLIC_PYTHON_BASELINE_ONLY`

- 현재 공개 기준선은 `ownership/reuse-reconciliation.v1.json`에 고정된 Python package, CLI와 schema입니다.
- `dyarchy-pipeline/src/hwp2yaml-rs/**`는 동일 도메인의 후보 구현일 뿐이며 이 저장소의 source, release 또는 compatibility authority가 아닙니다.
- Rust package·binary·schema를 이 저장소에 넣는 결정은 별도 ADR, provenance/license 승인, standalone build와 synthetic conformance evidence 없이는 금지합니다.

`REUSE_IS_NOT_CUTOVER`

- 이 저장소가 재사용 대상으로 선택되었다는 사실은 source import, compatibility, release, migration, cutover 또는 원본 삭제 승인이 아닙니다.
- Python CLI, Rust CLI, in-process Rust port와 pipeline subprocess consumer 계약을 각각 검증하고 단일 package·CLI·schema·semantic-version authority를 확정해야 합니다.

`NO_DUPLICATE_REMOTE`

- 같은 제품을 위한 `hwp2yaml-rs`, `dyarchy-hwp2yaml` 또는 유사 중복 원격을 만들지 않습니다.
- 기능이 이 경계에 맞지 않으면 강제로 혼입하지 않고 ownership manifest에 근거 있는 HOLD로 남깁니다.

## 공개 저장소 개인정보와 provenance

`PUBLIC_DATA_FAIL_CLOSED`

- 실제 고객·시험 응시자·분쟁·보험·의료·금융 문서, raw HWP/HWPX/PDF/OCR, 변환 전문, 계정·사건 식별자, credential, private key, host 절대경로와 live runtime snapshot을 commit하지 않습니다.
- fixture는 최소 synthetic data만 허용합니다. 공개 자료를 사용할 때도 원본 URL·고정 revision·재배포 근거·비식별 검토·최소 데이터 필요성을 machine-readable manifest에 남겨야 합니다.
- privacy 또는 출처가 불명확한 기존 fixture는 내용 재사용·복사·공개 이력 import 없이 `REVIEW_REQUIRED`로 격리합니다.

`IMPORT_PROVENANCE_REQUIRED`

- 외부 source import에는 source repository, exact commit/release, original path, exact blob, author/copyright, SPDX license, NOTICE 의무, 변경·생성 recipe와 승인자를 기록합니다.
- 자동 생성 table도 generator, 입력 권위, revision, license와 재현 hash가 없으면 source import할 수 없습니다.
- import는 immutable Git object에서만 수행하며 dirty working tree나 임시 산출물 복사를 금지합니다.

`NO_SOURCE_DELETION_WITHOUT_APPROVAL`

- 원본 `dyarchy-pipeline`, Dya, ROVER, branch, worktree, remote 또는 Syncthing source는 이 저장소가 존재하거나 재사용 대상으로 선택되었다는 이유로 삭제·이동·archive·rename하지 않습니다.

## 비신뢰 문서와 외부 실행 경계

`UNTRUSTED_INPUT_LIMITS_REQUIRED`

- HWP/HWPX/ZIP/XML/OLE/DEFLATE 입력은 비신뢰 입력으로 취급합니다.
- input bytes, archive entry 수, entry·총 압축해제 bytes, 압축비, XML depth/events/attributes, stream/record/section/paragraph/table/cell 수, text/output bytes, wall time, memory와 concurrency에 명시적 상한을 둡니다.
- 모든 크기 산술과 allocation은 overflow·남은 입력 길이·quota를 먼저 검사하고, 초과는 panic/OOM/부분 성공이 아닌 안정된 typed error로 종료합니다.

`ARCHIVE_XML_FAIL_CLOSED`

- absolute path, drive prefix, `..`, NUL 또는 비정상 archive member name을 거부합니다.
- DTD, external entity, entity expansion, malformed XML/OLE/ZIP, encrypted 또는 unsupported 문서는 output을 만들지 않고 fail closed해야 합니다.

`EXTERNAL_PROCESS_OPT_IN`

- LibreOffice, Docling, pdftotext, Python/model runner와 network/model download는 기본 parser 성공의 암묵적 fallback이 될 수 없습니다.
- 외부 실행은 명시적 capability, 고정 executable/version, argument-array 호출, 호출별 격리 tempdir, no-egress 기본값, timeout 시 process-tree 종료와 artifact 정리를 가져야 합니다.
- 도구 부재·실패·timeout을 success나 partial YAML로 승격하지 않습니다.

## 결정적 출력과 제품 주장

`DETERMINISTIC_OUTPUT_REQUIRED`

- schema version, field 의미, ordering, newline과 error semantics를 고정합니다.
- clock, hostname, user name, 절대 source/temp path와 file URI는 출력에서 제거하거나 명시적으로 주입된 deterministic 값으로 바꿉니다.
- 동일 input/options를 서로 다른 clean path에서 반복 실행했을 때 canonical semantic result가 같아야 합니다.

`SINGLE_VERSION_AUTHORITY`

- package metadata, importable `__version__`, CLI `--version`, artifact와 release tag는 하나의 권위에서 생성되어 일치해야 합니다.
- 문서의 CLI 예시와 실제 parser command·option·exit code가 다르면 release할 수 없습니다.

`CLAIMS_REQUIRE_EVIDENCE`

- 지원 버전, 완전 보존, 보안, 성능 또는 호환성 주장은 공개 synthetic fixture matrix와 exact-head CI evidence로 뒷받침해야 합니다.
- detection, direct parsing, external-tool conversion과 layout enrichment를 하나의 “지원” 상태로 합치지 않습니다.

## 하네스 진화

`HARNESS_EVOLUTION_REQUIRED`

- 새 parser failure, path/privacy leak, dependency advisory, license/provenance gap, CLI/schema drift 또는 nondeterminism이 발견되면 같은 변경에서 최소 synthetic fixture, negative canary, verifier, 문서와 CI를 함께 보강합니다.
- 실패한 test를 skip하거나 assertion, quota, privacy scan, provenance field 또는 supported-platform 조건을 삭제·완화해 통과시키지 않습니다.

`AGENTS_HARNESS_COHERENCE`

- 이 헌법의 machine-readable marker와 실제 verifier/workflow가 같은 계약을 검사해야 합니다.
- 문구만 존재하고 실행 evidence가 없는 완료 주장을 금지합니다.

`POSITIVE_AND_NEGATIVE_BEHAVIOR_REQUIRED`

- 승인된 HWP5/HWPX/HWP3 surface에는 실제 성공 synthetic case가 필요합니다.
- malformed, encrypted, oversized, ZIP bomb, unsafe member, entity attack, external-tool failure, source hash drift, forbidden corpus와 path leak은 명시적으로 거부되어야 합니다.

## branch와 worktree

`ROLE_SEPARATED_WORKTREES`

- `main`: 현재 공개 Python 제품과 검증된 공통 거버넌스만 보유합니다.
- `governance/hwp-reuse-reconciliation`: 헌법, ownership, provenance, include/exclude와 lifecycle 문서만 다룹니다.
- `verify/hwp-reuse-contract`: verifier, synthetic fixture, workflow와 evidence만 다룹니다.
- `feature/hwp-contract-alignment`: 승인된 API/CLI/schema/security 구현 변경만 다룹니다.
- `migration/rust-implementation-import`: 별도 승인된 immutable source history와 import mapping만 다룹니다.
- 각 역할 branch는 별도 sibling worktree에서 작업합니다.

부모 통합 담당자만 공통 파일 merge, 충돌 해결, PR 상태 변경과 최종 correctness/security/release 판단을 소유합니다.

## 필수 검증

`SUPPORTED_HOSTED_MATRIX`

- 현재 최소 hosted matrix는 `ubuntu-latest`와 `windows-latest` 각각의 CPython 3.10 및 3.12입니다.
- macOS와 Rust artifact는 현재 지원·release matrix가 아닙니다. 이를 추가하려면 machine-readable matrix, dependency/toolchain pin과 해당 hosted/fresh-clone evidence를 같은 변경에 추가합니다.

```text
python -B qa/verify_repository.py
python -B -m unittest discover -s qa -p "test_*.py"
python -m pytest
ruff check src tests qa
package/import/CLI version parity
synthetic HWP5/HWPX/HWP3 positive and adversarial matrix
dependency advisory, license, provenance and SBOM gates
git diff --check
git fsck --full --strict
```

Rust 구현을 승인해 들여오는 경우에만 다음을 추가합니다.

```text
cargo fmt --all -- --check
cargo metadata --locked
cargo check --locked --all-targets
cargo clippy --locked --all-targets -- -D warnings
cargo test --locked --all-targets
cargo audit and cargo deny license/source gates
offline fresh-clone build on each supported platform
```

`EXTERNAL_CI_AND_FRESH_CLONE_REQUIRED`

- merge 전에는 exact integration head의 지원 OS hosted jobs가 실제 생성되어 통과하고, 별도 fresh clone에서 바로 그 integration commit을 checkout한 전 gate 검증과 security/provenance review가 끝나야 합니다.
- merge 후에는 실제 final `main` commit을 새로 clone해 다시 검증하기 전 `RECONCILED`, `MIGRATED`, public release 또는 production cutover를 선언하지 않습니다.
- `startup_failure`, jobs 0, draft review 생략, clean mergeability, local-only PASS 또는 기존 이름의 동일성은 충분하지 않습니다.
