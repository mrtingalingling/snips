---
name: test-and-ci
description: Audits and then, with approval, adds CI/CD pipelines, unit tests, smoke tests, regression tests, dependency audits, multi-device testing and slow-network testing to an existing app, in any stack. Use when someone wants to set up CI, GitHub Actions or another pipeline, add or improve tests, add post-deploy smoke checks, stop bugs from coming back, audit dependencies for vulnerabilities, check how the app behaves on phones or tablets, or test on 3G or flaky connections — even if they only say "make sure this doesn't break" or "add some tests before launch". Part of the ship-readiness family (items T1–T7).
---

# Test and CI

Tests and CI exist so that a bad change gets caught by a machine before a user catches it. This skill audits what exists, proposes the smallest set of additions that closes the real gaps, and implements only what the user approves.

## Workflow

1. **Detect the stack.** Read the manifests (package.json, pyproject, Gemfile, go.mod, Cargo.toml…), existing test config, CI config, and the hosting config. If the `ship-readiness-audit` skill is installed (workspace `.agents/skills/` or global `~/.gemini/config/skills/`), its `scripts/triage_scan.py` does this in one pass (run it with `--help` first). Note the package manager and lockfile. CI must use the same ones.
2. **Audit T1–T7** below. Record each as PASS / PARTIAL / MISSING / UNKNOWN / N/A, with evidence (file path and line, or command output). Run the existing test suite and record the actual result. "Tests exist" and "tests pass" are different findings.
3. **Report and stop.** Use the ship-readiness report table (ID, Item, Status, Severity, Evidence, Fix, Size, Who), then ask which items to implement. Don't add tests or pipelines before approval.
4. **Implement approved items** one per branch or commit. If `reviewable-diffs` is installed, follow it: the approved items are the requirements. Check current CLI flags, action versions and config keys in the tool's own docs (`live-source-verification`) rather than writing them from memory. CI config written from memory is the most common source of pipelines that fail on the first run.
5. **Verify.** Run every new test locally and confirm it fails when the behavior is broken (temporarily break it, or reason explicitly about why it would fail). For CI changes in repo mode, validate the config syntax with the tool's linter if one exists. Say plainly that the pipeline hasn't actually run until it runs on the CI provider.

Without workspace or terminal access, you can still produce the pipeline file, test files and runbooks, but mark them "unrun" and give the user the exact commands to run them.

## T1 — CI/CD

Good looks like: on every PR, install from the lockfile, then typecheck/lint, unit tests, build. On merge to main: the same, then deploy, then smoke tests (T3) against the deployed URL. Deploy depends on the earlier jobs passing. Secrets come from the CI secret store, never the repo.

Audit points:
- Does deploy happen outside CI (manual pushes from a laptop, host auto-deploys on every push regardless of test results)? Hosts that auto-deploy from git often deploy even when CI fails. Either gate it (required status checks, branch protection, or deploy-from-CI only) or report it as PARTIAL.
- Is the install reproducible (lockfile-respecting install command, pinned runtime version)?
- Are third-party CI actions/orbs pinned to a version or commit, not a moving branch?
- Is there caching for dependencies? This is a speed matter, LOW severity.

Branch protection and required checks are repo settings, so they're a Human task. Write the exact settings to enable.

## T2 — Unit tests

Good looks like: the logic that would hurt if wrong is covered. That means money, permissions, data transformations, parsing, state machines and date/time handling. Coverage percentage is a weak proxy. Ask "which function, if broken, would cause a support ticket or a refund?" and test those first.

When adding tests to an untested codebase, don't try to reach a coverage number. Propose 5–15 high-value tests, list them with one-line rationales in the report, and let the user pick. Use the test framework already present. If none, choose the one most idiomatic for the stack and say it's a preference, not a requirement.

## T3 — Smoke tests

Smoke tests answer one question after a deploy: is the app basically alive? They should be few (3–8), fast (under a couple of minutes), and run against the real deployed URL.

Typical set: home page returns 200 and contains an expected string, health endpoint returns OK (coordinate with `ops-reliability` O5), login page renders, one authenticated core action using a dedicated test account, and one API call. Use a test account whose credentials live in CI secrets. Never use a real user's account.

Wire failure to consequences. The deploy is marked failed, a human is notified, and if the host supports it, promotion is blocked or rollback is triggered (`ops-reliability` O7). A smoke test that fails silently is decoration.

Don't run destructive actions (payments, emails to real addresses, deletes) against production in smoke tests. Use staging for those, or provider test modes.

## T4 — Regression tests

A regression test pins a bug that already happened so it can't come back. The practice: reproduce the bug in a failing test first, then fix, and keep the test named after the bug (issue number or one-line description).

For the audit, look at recent fix commits and check whether each brought a test. If the user can list past bugs, offer to backfill tests for the ones most likely to recur. Visual regression (screenshot diffing) is optional and adds maintenance. Suggest it only for UI-heavy apps where layout breakage has actually happened.

## T5 — Dependency audit

Use each ecosystem's own audit tool, for example `npm audit` / `pnpm audit` / `yarn npm audit`, `pip-audit`, `bundler-audit`, `cargo audit`, `govulncheck`, or the cross-ecosystem `osv-scanner`. Confirm current flags before scripting them.

Triage before fixing. Most raw audit output is noise:
- **Production vs dev-only.** A vulnerability in a build tool is usually lower risk than one in a request handler.
- **Reachability.** Is the vulnerable function actually used? Say how you checked, or mark it UNKNOWN.
- **Fix available?** Prefer the minimal version bump that resolves it. Each bump gets its own commit with the advisory ID in the message, so a breakage can be bisected.
- Avoid blanket "fix all" commands that rewrite the lockfile wholesale. They produce unreviewable diffs and surprise major upgrades.

Add the audit to CI at a severity threshold the user chooses. Failing on every low-severity advisory trains people to ignore the job.

## T6 — Multi-device testing

What can be automated: run e2e tests under device emulation profiles (viewport, user agent, touch, device scale). Use at least one mobile WebKit profile (closest to iOS Safari), one mobile Chromium profile, and desktop. Check small-width layout (around 320–375px CSS width), tap targets, horizontal overflow, and that nothing depends on hover.

What emulation can't prove: real iOS Safari quirks (viewport units with the toolbar, input zoom, safe areas), real performance on low-end phones, and real keyboards. Say so in the report. Make a short real-device pass a Human task before launch, either on the user's own phone plus one other or via a device cloud. Write the checklist they should follow.

## T7 — Slow-network testing

Goal: the app stays usable and honest on slow or flaky connections. Check:
- Loading states appear quickly and don't flash (coordinate with `launch-polish` P4).
- Buttons that submit are disabled or deduplicated while in flight. Double-submits on slow networks cause duplicate orders and duplicate sign-ups.
- Requests have timeouts and show a retry path instead of spinning forever.
- Initial page weight is reasonable. Report JS and image bytes for the main route from the build output.

How: browser devtools presets for a manual pass, or programmatic network throttling in the e2e tool. In some tools this is only available for Chromium, so verify against current docs. Also test offline-then-online for any flow that saves data. Write the throttled scenarios as e2e tests where the tool supports it, otherwise as a short manual script in the repo (`docs/manual-tests.md`).
