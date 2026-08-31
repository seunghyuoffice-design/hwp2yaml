# Repository ownership and HWP reuse reconciliation

## Current decision

`hwp2yaml` is the only selected reuse target for the HWP/HWPX conversion domain. No duplicate successor repository is authorized.

This is not a completed reuse or cutover. The current lifecycle is `HOLD_DEPENDENCY_RECONCILE` because the public Python implementation, the pipeline Rust implementation and the pipeline runtime consumer do not expose the same CLI, schema or dependency boundary.

No product source was moved or changed by this receipt. The dirty Dya working tree is not an import source, and neither the pipeline source nor its history may be deleted, moved, archived or renamed without separate explicit approval.

## Immutable baselines

The machine-readable receipt implementing the higher-level `AGENTS.md` constitution is `ownership/reuse-reconciliation.v1.json`.

- Public target: `seunghyuoffice-design/hwp2yaml@56231e2e4e26e3aff3dcff6ef82f5789aaa4a9d8`, root tree `0663b1cdb86d2b73dae1c25433a0f9ddefd06c1d`, 21 tracked files.
- Pipeline Rust candidate: `seunghyuoffice-design/dyarchy-pipeline@b565aa04737389e21a4113a2ec2af0ca1fb68c76`, component tree `2011ab9bad0c97941622044b07c8501073c7b708`, 33 paths under `src/hwp2yaml-rs/`.
- Dya-v6 legacy evidence: `seunghyuoffice-design/Dyarchy-v6@52f80a13631e3054b76c4f9fd0df7a8c28161176`, component tree `8e5733b817bd8085e8c8bf2f9fc539ae9163bf82`, 32 paths under `archive/Dyarchy-v4/src/hwp2yaml-rs/`.
- The pipeline import receipt maps the Rust scope to `seunghyuoffice-design/Dyarchy-v4@6d29588e274255a5e1f9752ac21cc5c5311baa05`; its first pipeline scope commit is `ebe6104f61170e1f29387ea2f6af14be25b665fc`.

Only immutable Git objects were used to calculate these baselines. The Dya working tree and untracked data were excluded.

## Contract mismatch receipt

| Surface | Observed command or contract | Current decision |
|---|---|---|
| Public Python CLI | `extract`, `batch`, `info` | Current repository release surface only; future authority not yet decided |
| Pipeline Rust main CLI | `convert`, `batch`, `info`, `text` | Candidate evidence only; not imported |
| Pipeline runtime consumer | `parse --input ... --output-format yaml` | No matching command in either observed main CLI; integration blocked |
| In-process Rust bridge | `conversion-port::DocumentConverter` and `default_converter()` | Pipeline-relative dependency; not standalone in this repository |

Additional baseline drift is intentionally visible rather than hidden:

- `pyproject.toml` declares `0.6.1`, while `src/hwp2yaml/__init__.py` declares `0.6.0`.
- README CLI examples omit the subcommand required by the current parser.
- A local-only Windows smoke at the exact public baseline passed 26 pytest cases and `pip check`, while Ruff reported 44 findings. The command, tool versions and output digests are sealed in the manifest. This is not hosted or reproducible release evidence because the repository has no dependency lockfile.
- Python and Rust outputs contain different models and volatile metadata paths; semantic and byte determinism are unproved.
- Actual HWP3 conversion depends on external tools that were absent from the local audit host.

These are recorded blockers, not release evidence.

## Public-import stop conditions

The Rust candidate must not be copied, history-filtered or pushed into this public repository yet.

Two immediate public-import blockers are sealed:

1. Third-party-derived code/data in the candidate lacks a complete source revision, copyright, SPDX license, NOTICE and generation chain suitable for public redistribution.
2. `src/hwp2yaml-rs/test_files/1999-046.yaml` has no approved public provenance/privacy receipt and is not an allowed successor fixture. Its content and history are excluded from public import.

Further blockers include the pipeline-relative `conversion-port` dependency, implicit external executable/model boundaries, missing parser resource quotas, cancellation gaps, known dependency-advisory review needs, unsafe long-Unicode handling and unresolved package/CLI/schema authority.

## Included evidence scope

- The exact 21-file public Python baseline at the pinned target commit.
- The exact pipeline Rust component tree only as immutable comparison evidence.
- The exact Dya-v6 archive component tree only as immutable legacy comparison evidence.
- Eight pinned pipeline consumer/configuration paths listed in the reconciliation manifest.
- Future, minimal, generator-backed synthetic HWP3/HWP5/HWPX fixtures and their hashes.

## Excluded scope

- All raw or real HWP/HWPX/PDF/OCR documents and converted document bodies.
- The unresolved Rust YAML fixture and any private/customer/dispute/insurance/medical/financial corpus.
- Document normalizer, DOCX, PDF/XML, layout, dashboard, quality, Redis, queue, watchdog, power, SSH, Docker and deployment code.
- Dya or pipeline dirty working-tree files, untracked files, host paths, credentials and live runtime state.
- Rust source/history import, CLI compatibility shims and product source edits until their role-specific branches are separately approved.

## Required transition

1. Keep the governance PR draft and unmerged while a separate sibling verify worktree adds the fail-closed verifier, negative canaries and hosted workflow on top of the unchanged governance commit.
2. The parent integrator must present one final integration head containing both role-separated commits. That exact head, not the governance-only head, must create and pass every required hosted job before either commit reaches `main`.
3. Before merge, a separate fresh clone must checkout that same exact integration commit and pass the full available governance/product gate. The hosted-CI head, requested checkout and observed checkout must be identical.
4. Seal the immutable receipt, forbidden-corpus scan and contract mismatch with positive and negative tests, then independently review the final integration diff.
5. After an approved merge, fresh-clone the resulting actual `main` commit and repeat the gates before calling this repository reconciled or using it for release/cutover.
6. Decide a single package, CLI, schema and version authority in an ADR.
7. Resolve all public provenance/license and fixture privacy blockers before any Rust history enters this public repository.
8. Build deterministic, bounded, synthetic conformance and external-tool failure gates.
9. Only then open a source-only feature or migration branch; the same pre-merge integration and post-merge final-main gates apply again.
