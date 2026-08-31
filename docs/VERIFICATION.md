# Repository verification

## Current evidence level

The verification branch is stacked on governance commit `6f6ebd8ce9cfe27e6331a3afa3e38971fe496e45`. It adds only verifier, negative-canary, Ruff regression-floor, workflow and verification-document paths. Product source and tests remain byte-identical to public baseline `56231e2e4e26e3aff3dcff6ef82f5789aaa4a9d8`.

The governance receipt intentionally remains `HOLD_DEPENDENCY_RECONCILE`. Adding a local verifier does not approve Rust import, CLI/schema compatibility, source movement, merge, release or cutover.

## Local gates

```text
python -B qa/verify_repository.py
python -B -m unittest discover -s qa -p "test_*.py"
python -B -m pytest -q -p no:cacheprovider tests
python -B qa/verify_ruff_baseline.py
python -m pip check
git diff --check 56231e2e4e26e3aff3dcff6ef82f5789aaa4a9d8 --
git fsck --full --strict
```

`qa/verify_repository.py` verifies:

- the exact 21-file public Python baseline and unchanged product paths;
- the sealed governance file blobs and required constitution markers;
- HOLD lifecycle, false approvals, no source mutation, exact consumer pins and both fresh-clone gates;
- the exact additive governance/verify path set;
- current and reachable-history rejection of raw PDF/OLE-HWP/HWP3/ZIP-HWPX signatures even after renaming, private/archive paths, large or non-tree blobs, secret-like values in content/path/commit/tag metadata, personal absolute paths and file URIs.

The verifier treats pipeline and Dya commit/tree/blob values as a sealed external receipt. It cannot prove those external objects exist from a fresh `hwp2yaml` clone; that comparison requires the independently retained pipeline and Dya object stores. It also cannot prove that GitHub private vulnerability reporting remains enabled or that a remote PR, hosted job, review or merge exists.

## Ruff debt is not a pass

The public baseline has 44 Ruff findings. `qa/verify_ruff_baseline.py` pins Ruff 0.16.5 and the normalized finding fingerprint so the debt cannot silently grow, move or be relabeled. A passing regression-floor command means only “the known debt is unchanged.” It is not a clean Ruff result and cannot satisfy the release gate. A source feature must reduce the findings to zero and replace this HOLD receipt with an exact clean run.

The raw-document allowlist is empty. Plain-text OCR has no reliable universal magic, so no future corpus or fixture path may be added merely because content sniffing returns unknown. Any future synthetic fixture requires a generator, hash, provenance/privacy receipt, explicit allowlist update and review; the exact additive-path verifier otherwise rejects it.

## Hosted and fresh-clone gates

The workflow requires `ubuntu-latest` and `windows-latest`, each with CPython 3.10 and 3.12. Action references are pinned to exact Git commits.

The following remain unverified until live GitHub evidence is attached:

- every expected matrix job is actually created and succeeds at one exact integration head;
- an independent pre-merge fresh clone checks out that identical integration commit and passes the gates;
- governance, verification and security review approvals are dateable and attached;
- after merge, a new clone observes the actual final `main` commit and repeats the gates.

`startup_failure`, jobs `[]`, local PASS, a successful Ruff regression floor, clean mergeability or a draft PR are not hosted-CI or release evidence.

## Re-run triggers

Re-run all gates when any governance, verifier, workflow, baseline source/test, dependency metadata, supported matrix, provenance pin, privacy rule, CLI/schema claim or fresh-clone evidence changes. Parser/source changes additionally require the synthetic positive/adversarial, dependency advisory, license, SBOM and external-process gates in `AGENTS.md`.
