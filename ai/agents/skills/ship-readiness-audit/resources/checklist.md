# Ship-readiness checklist (40 items)

Each item lists where the evidence lives (**repo**, **external**, or **both**), what PASS looks like, how to detect it, and a one-line fix direction. Depth for each area lives in the sibling skill named in the heading.

## Contents
- T — Testing and CI (`test-and-ci`): T1–T7
- S — Security and auth (`security-hardening`): S1–S7
- O — Operations and reliability (`ops-reliability`): O1–O11
- A — LLM features (`llm-hardening`): A1–A4
- P — Launch polish and compliance (`launch-polish`): P1–P11

---

## T — Testing and CI → `test-and-ci`

**T1 CI/CD** · repo
PASS: a pipeline config (e.g. `.github/workflows/`, `.gitlab-ci.yml`, `circle.yml`, `azure-pipelines.yml`) runs install from lockfile, typecheck/lint, tests and build on every PR, and deploys from the main branch only after those pass.
Detect: read the CI config. Check triggers, job order and whether deploy depends on test jobs.
Fix: add a pipeline with those stages. Keep secrets in the CI secret store.

**T2 Unit tests** · repo
PASS: a test runner is configured, tests exist for core business logic (not just a sample test), and they run in CI.
Detect: test framework in manifest, test files, a `test` script. Run it and record the result.
Fix: cover pure logic first (pricing, permissions, parsing, state transitions).

**T3 Smoke tests** · repo
PASS: a small suite hits critical paths (home, login, one core action, health endpoint) against a deployed URL after each deploy.
Detect: e2e config (Playwright/Cypress/etc.) plus a CI step that runs it against a preview or production URL.
Fix: 3–8 fast checks, run post-deploy, and failure blocks promotion or triggers rollback.

**T4 Regression tests** · repo
PASS: fixed bugs get a test that reproduces them. The full suite runs in CI.
Detect: commit history of fixes vs. tests added alongside. A test name or comment referencing the bug.
Fix: adopt "no bug fix without a failing test first". Add tests for known past bugs.

**T5 Dependency audit** · repo
PASS: the ecosystem's audit tool runs in CI. No unaddressed high/critical advisories in production dependencies.
Detect: run the audit tool for each detected ecosystem and record output.
Fix: triage by reachability. Do minimal bumps in their own branch. Add the audit to CI.

**T6 Multi-device testing** · both
PASS: e2e tests run in at least one mobile WebKit and one mobile Chromium profile plus desktop. Layout checked at small widths.
Detect: e2e config projects/devices.
Fix: add device projects. Real-device check before launch is a Human task.

**T7 Slow-network testing** · both
PASS: key flows tested under throttled network. Loading states appear, requests time out sensibly, no duplicate submits.
Detect: throttling in e2e setup, or a documented manual pass.
Fix: add a throttled e2e project or a manual test script.

## S — Security and auth → `security-hardening`

**S1 Auth on every endpoint** · repo
PASS: every route, API handler, server action, RPC, GraphQL resolver and webhook has a deliberate rule: public, authenticated, or authorized-for-resource. Tests prove unauthenticated → 401 and wrong user → 403/404.
Detect: enumerate routes and build a route matrix. Look for default-deny middleware and per-resource ownership checks.
Fix: default-deny, explicit public allowlist, ownership checks, signed webhooks.

**S2 Secrets out of the repo** · repo
PASS: no secrets in tracked files or history. `.env*` ignored (except an example file with placeholders). No secrets in client bundles.
Detect: triage scanner hits, `git ls-files | grep -i env`, history scanner (gitleaks/trufflehog or similar), public env-var prefixes.
Fix: move to secret store and env vars. Add ignore rules and a pre-commit secret scan.

**S3 Rotate leaked keys** · external
PASS: every secret ever committed or exposed has been rotated at the provider and the old one revoked.
Detect: S2 findings, including history. Ask the user to confirm rotation per key.
Fix: Human runbook: rotate, update stores, redeploy, revoke old, check provider logs.

**S4 Rate limiting** · repo
PASS: limits on auth, signup, password reset, verification resend, and expensive endpoints (LLM, uploads, email-sending). The limiter store is shared across instances. Returns 429 with Retry-After.
Detect: limiter middleware/library, its store, which routes use it.
Fix: per-IP plus per-account limits with a shared store.

**S5 Password reset works** · both
PASS: single-use, expiring, hashed-at-rest tokens. No account-existence leak. Link built from configured base URL. Sessions invalidated after reset. Tested end-to-end with a real inbox.
Detect: read the flow. Run it in staging.
Fix: per security-hardening S5. The real-inbox test is a Human or staging task.

**S6 Email verification works** · both
PASS: verification required where it matters, tokens expire, resend is rate-limited, changing email re-verifies. Tested end-to-end.
Detect: read the flow. Run it in staging.
Fix: per security-hardening S6.

**S7 SPF / DKIM / DMARC** · external
PASS: sending domain has one SPF record covering all senders, DKIM for each provider, and a DMARC record with reporting.
Detect: `dig TXT <domain>`, `dig TXT _dmarc.<domain>`, provider's DKIM selector. N/A if the app sends no email.
Fix: Human DNS edits from the provider's current docs. Start DMARC at `p=none` with reports, then tighten.

## O — Operations and reliability → `ops-reliability`

**O1 Logging** · repo
PASS: structured logs (JSON or key=value) with request IDs, levels, no secrets or PII in payloads, shipped somewhere searchable with retention.
Detect: logger library and config, log calls in request handlers, log drain config.
Fix: structured logger, request-ID middleware, redaction.

**O2 Error tracking** · both
PASS: an error tracker captures server and client errors with release/version tags and source maps where relevant.
Detect: SDK in dependencies and initialized at entrypoints. DSN from env, not hardcoded.
Fix: install and init on server and client. Tag releases in CI.

**O3 Alerting** · external
PASS: alerts fire to a human channel for error-rate spikes, downtime, failed jobs and spend thresholds, with an owner.
Detect: ask. Alert rules may be in code (IaC) or dashboards.
Fix: define 3–5 alerts with thresholds. Human configures channels.

**O4 Performance monitoring** · both
PASS: server latency/throughput tracked. Real-user web vitals collected for the web front end.
Detect: APM/tracing SDK, web-vitals reporting, host-provided analytics.
Fix: enable tracing in the error tracker or APM. Report web vitals.

**O5 Uptime monitoring** · external
PASS: an external monitor checks a health endpoint and the home page from outside the host, at an interval suited to the app, alerting a human.
Detect: ask. Check for a `/health` endpoint in the repo.
Fix: add a health endpoint that checks dependencies. Human sets up the monitor.

**O6 Staging environment** · both
PASS: a separate environment with its own database and secrets, deployed from a branch or PR previews, that mirrors production config.
Detect: CI environments, host config, separate env files or env-scoped secrets.
Fix: per ops-reliability O6. Staging must not touch production data or send real emails/charges.

**O7 Rollback plan** · both
PASS: a documented, practiced, one-step way to return to the previous release, with migrations that tolerate rollback (expand/contract).
Detect: `ROLLBACK.md` or similar, host's deploy history feature, migration style.
Fix: write the runbook, make migrations backward compatible, rehearse once in staging.

**O8 Database backups** · external
PASS: automated backups (and point-in-time recovery if available) with known retention, stored separately from the primary.
Detect: ask, or read IaC/provider config.
Fix: Human enables in provider. Document schedule and retention.

**O9 Restore actually tested** · external
PASS: a backup was restored to a scratch database within a stated recent period, verified by checks, and the steps and time taken are recorded.
Detect: ask for the date of the last restore test and the runbook.
Fix: per ops-reliability O9. Write `RESTORE.md`, do a timed drill.

**O10 Caching** · repo
PASS: static assets are fingerprinted with long cache lifetimes. HTML and API responses have deliberate cache headers. No user-specific data cached publicly.
Detect: response headers (`curl -I`), framework cache config, CDN config.
Fix: per ops-reliability O10.

**O11 Spend caps and billing alerts** · external
PASS: every pay-per-use provider (cloud, LLM, email, SMS, maps) has a budget alert and, where offered, a hard cap. In-app per-user quotas protect pay-per-use endpoints.
Detect: list providers from dependencies and env vars. Ask about each dashboard.
Fix: Human sets caps and alerts per provider. The agent adds in-app quotas.

## A — LLM features → `llm-hardening` (N/A if the app makes no model calls)

**A1 Evals** · repo
PASS: a versioned set of representative inputs with graders (exact, rubric, or model-graded), run before prompt or model changes and ideally in CI.
Detect: eval directory, scripts, CI step.
Fix: per llm-hardening A1.

**A2 Prompt-injection testing** · repo
PASS: untrusted content (user input, retrieved docs, web pages, tool output, uploads) is tested for instruction-following. Tools are least-privilege. Model output is never executed or rendered as raw HTML.
Detect: test cases for injection, tool permission design, output handling.
Fix: per llm-hardening A2.

**A3 Model API fallback** · repo
PASS: timeouts, retries with backoff and jitter on retryable errors, a fallback path (other model/provider, cached answer, or graceful message), and user-visible degraded state.
Detect: API client wrapper, retry config, error handling around calls.
Fix: per llm-hardening A3.

**A4 Token limits and context overflow** · repo
PASS: input size is measured or bounded before sending. There is a strategy when it exceeds the budget (truncate, summarize, retrieve). Output truncation (max-tokens stop) is detected and handled.
Detect: token counting, history trimming, handling of the stop reason.
Fix: per llm-hardening A4.

## P — Launch polish and compliance → `launch-polish`

**P1 SEO and AEO** · repo
PASS: `robots.txt` allows production pages and references the sitemap. Staging is blocked from indexing. Valid `sitemap.xml`. Unique titles and meta descriptions. Canonical URLs. Structured data (JSON-LD) that validates. Key content renders without client JS.
Detect: fetch or read these files. View server-rendered HTML.
Fix: per launch-polish P1.

**P2 OG images and favicon** · repo
PASS: Open Graph and Twitter card tags with an image on shareable pages. Favicon set including apple-touch-icon.
Detect: inspect `<head>`.
Fix: add tags and assets.

**P3 Image optimization** · repo
PASS: modern formats where supported, responsive sizes, explicit width/height (no layout shift), lazy-loading below the fold, no multi-MB originals served.
Detect: image components/tags, asset sizes in build output.
Fix: per launch-polish P3.

**P4 404, 500, empty and loading states** · repo
PASS: custom 404 and 500 pages with a way back. Every data view has designed empty, loading and error states.
Detect: framework error pages, components with data fetching.
Fix: per launch-polish P4.

**P5 Accessibility** · both
PASS: automated scan (e.g. axe) clean on key pages. Keyboard-only pass works. Visible focus, labels, alt text, contrast. Target WCAG 2.2 AA.
Detect: run an automated scanner. Manual keyboard pass is partly Human.
Fix: per launch-polish P5.

**P6 Analytics** · both
PASS: a tracking tool records page views and 3–5 named key events tied to the product's goals, respecting consent (P9).
Detect: analytics SDK, event calls.
Fix: per launch-polish P6.

**P7 Custom domain and SSL** · external
PASS: production on the custom domain over HTTPS, HTTP redirects to HTTPS, one canonical host (apex or www), certificate auto-renews.
Detect: `curl -I http://<domain>`, `curl -I https://<domain>`.
Fix: Human DNS and host settings, plus redirect config.

**P8 Privacy policy and terms** · both
PASS: pages exist, are linked in the footer and at signup, and accurately describe the data actually collected and the processors actually used.
Detect: routes or pages. Compare to the inventory of SDKs and providers.
Fix: draft from an accurate data inventory, marked for legal review. Not legal advice.

**P9 Cookie consent** · both
PASS: if non-essential cookies or trackers are used where consent is required, they don't load until consent. Choice is stored and changeable.
Detect: cookie/tracker inventory, consent tool, whether scripts are gated.
Fix: per launch-polish P9. Jurisdiction-dependent.

**P10 Account and data deletion** · repo
PASS: users can delete their account in-product. Deletion covers the primary DB, files, auth provider and third-party processors. Backup retention is documented.
Detect: delete-account flow and what it touches.
Fix: per launch-polish P10.

**P11 README** · repo
PASS: what it is, how to run locally, env vars (names only), how to test, how to deploy, how to roll back and restore, where monitoring lives.
Detect: read it.
Fix: write or extend it. Link the runbooks from O7 and O9.
