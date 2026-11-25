---
name: test-and-ci
description: Audits and then, with approval, adds CI/CD pipelines, unit tests, smoke tests, regression tests, dependency audits, multi-device testing and slow-network testing to an existing app, in any stack. Use when someone wants to set up CI, GitHub Actions or another pipeline, add or improve tests, add post-deploy smoke checks, stop bugs from coming back, audit dependencies for vulnerabilities, check how the app behaves on phones or tablets, or test on 3G or flaky connections — even if they only say "make sure this doesn't break" or "add some tests before launch". Part of the ship-readiness family (items T1–T7).
---

# Test and CI

Tests and CI exist so that a bad change gets caught by a machine before a user catches it. This skill audits what exists, proposes the smallest set of additions that closes the real gaps, and implements only what the user approves.

## Workflow

1. **Detect the stack.** Read the manifests (package.json, pyproject, Gemfile, go.mod, Cargo.toml…), existing test config, CI config, and the hosting config. If the `ship-readiness-audit` skill is installed (workspace `.agents/skills/` or global `~/.gemini/config/skills/`), the `triage_scan.py` script in its `scripts/` folder (run it with `--help` first) does this in one pass. Note the package manager and lockfile. CI must use the same ones.
2. **Audit T1–T7** below. Record each as PASS / PARTIAL / MISSING / UNKNOWN / N/A, with evidence (file path and line, or command output). Run the existing test suite and record the actual result. "Tests exist" and "tests pass" are different findings.
3. **Report and stop.** Use the ship-readiness report table (ID, Item, Status, Severity, Evidence, Fix, Size, Who), then ask which items to implement. Don't add tests or pipelines before approval. **Skip this step for items the user already approved in this conversation** (through `/ship-check` or the audit's Phase 4): go straight to step 4.
4. **Implement approved items** one per branch or commit. If `reviewable-diffs` is installed, follow it: the approved items are the requirements. Check current CLI flags, action versions and config keys in the tool's own docs (`live-source-verification`) rather than writing them from memory. CI config written from memory is the most common source of pipelines that fail on the first run.
5. **Verify.** Run every new test locally and confirm it fails when the behavior is broken (temporarily break it, or reason explicitly about why it would fail). For CI changes in repo mode, validate the config syntax with the tool's linter if one exists. Say plainly that the pipeline hasn't actually run until it runs on the CI provider.

Without workspace or terminal access, you can still produce the pipeline file, test files and runbooks, but mark them "unrun" and give the user the exact commands to run them.

## Working rules

- **App type first.** Use the app profile from the audit if one exists, otherwise classify the project with `resources/app-profiles.md` in the `ship-readiness-audit` skill. Mark items N/A for the profile instead of reporting them MISSING, and use the profile's changed meaning where the table says so.
- **Change discipline.** Use `reviewable-diffs` if installed (workspace `.agents/skills/` or global `~/.gemini/config/skills/`). Otherwise: one approved item per branch or commit; change only the lines that item needs; no formatting, renaming or refactoring churn; no features nobody asked for; add a test that fails without the fix and passes with it; list each file touched and why.
- **Current facts.** Use `live-source-verification` if installed. Either way, apply the citation rule: any vendor-specific or time-sensitive fact (pricing, plan limits, config keys, CLI flags, SDK and error names, DNS values, legal or store rules) carries an inline [Title](URL) retrieved this session. No URLs from memory. If a source is unreachable, write "Source unavailable" and tag the claim [Unknown].
- **Confidence tags.** Tag each finding's evidence with exactly one of [Cited], [Certain], [Likely], [Guessing], [Unknown] ([Cited]: checked against a live source this session, with the link; [Certain]: seen directly this session; [Likely]: strong inference; [Guessing]: hypothesis only; [Unknown]: name the gap and what would close it). PASS requires [Certain] or [Cited].
- **Irreversible steps need their own yes**, even for approved items: rewriting git history, rotating or revoking live keys, DNS changes, production deploys, restoring or deleting data, billing changes.

## Release items (CLI and library projects)

For CLI and library profiles this skill also owns R1–R5 (versioning and changelog, release pipeline, bad-release recovery, install smoke test, package contents). Their PASS criteria are in `resources/app-profiles.md` in the `ship-readiness-audit` skill. Check each registry's current rules live and cite them.

## Item playbooks

Detailed guidance for each item is in `resources/items.md`. Read only the sections for the items in scope, plus the **Current requirements** section at the end, which holds externally verified standards and dates:

- T1 — CI/CD
- T2 — Unit tests
- T3 — Smoke tests
- T4 — Regression tests
- T5 — Dependency audit
- T6 — Multi-device testing
- T7 — Slow-network testing
