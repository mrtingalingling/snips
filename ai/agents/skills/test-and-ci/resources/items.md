# Test and CI: item playbooks

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

## Current requirements (verified November 2026)

These change. Before relying on one in a report or runbook, re-check the linked source and cite it.

- **T1 Action pinning.** In March 2025 the `tj-actions/changed-files` GitHub Action was compromised and exposed workflow secrets of repositories that used it ([Unit 42](https://unit42.paloaltonetworks.com/github-actions-supply-chain-attack/)). Pin every third-party action to a full commit SHA, not a tag. Set a least-privilege `permissions` block per workflow, and don't run untrusted fork code with secrets (watch `pull_request_target`).
- **T5 Dependency cooldown and install scripts.** Self-replicating npm worms (Shai-Hulud, from September 2025) published malicious versions of popular packages. CISA's advice included pinning to known-safe versions and rotating developer credentials ([CISA](https://www.cisa.gov/news-events/alerts/2025/09/23/widespread-supply-chain-compromise-impacting-npm-ecosystem)). Defenses:
  - Set a minimum release age. npm added `min-release-age` in CLI 11.10.0 ([Socket](https://socket.dev/blog/npm-introduces-minimumreleaseage-and-bulk-oidc-configuration)), pnpm has `minimumReleaseAge` ([pnpm](https://pnpm.io/supply-chain-security)), and Dependabot has a cooldown option ([Datadog Security Labs](https://securitylabs.datadoghq.com/articles/dependency-cooldowns.md)).
  - Block or allow-list install scripts.
  - Install from the lockfile in CI (`npm ci`, `pnpm install --frozen-lockfile`).
- **T5 Taxonomy.** OWASP Top 10:2025 adds A03 Software Supply Chain Failures, which covers dependencies, build systems and delivery pipelines ([OWASP](https://owasp.org/Top10/2025/0x00_2025-Introduction/)). Tag supply-chain findings with it.
