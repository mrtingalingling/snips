# App profiles: which items apply

Classify the project in Phase 0 before auditing. A project can match more than one profile (for example a web app with a public API, or a library with a docs site). Then take the union of the applicable items.

For every item a profile marks N/A, record `N/A — <profile>: <reason>` in the report. Don't report it as MISSING, and don't propose fixing it.

## Detecting the profile

| Profile | Typical signals |
|---|---|
| **Web app** | Pages or templates, a front-end framework, HTML served to browsers |
| **API / backend service** | HTTP handlers returning JSON/gRPC, no user-facing pages, OpenAPI/GraphQL schema |
| **Worker / data pipeline** | Queue consumers, cron/scheduled jobs, ETL scripts, no inbound HTTP besides health |
| **CLI tool** | Executable entry point (`bin`, `console_scripts`, `main` package), argument parser, distributed via a package registry or binary releases |
| **Library / SDK** | Published package with an import API, no deployable service |
| **Mobile app** | iOS/Android/React Native/Flutter project; usually paired with an API profile for its backend |

If the signals conflict or are thin, ask the user one question to settle it. Don't guess.

## Applicability by profile

Legend: **✓** applies as written · **≈** applies with a changed meaning (see notes) · **–** N/A

| ID | Item | Web | API | Worker | CLI | Library | Mobile |
|---|---|---|---|---|---|---|---|
| T1 | CI/CD | ✓ | ✓ | ✓ | ≈ R2 | ≈ R2 | ✓ |
| T2 | Unit tests | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| T3 | Smoke tests | ✓ | ≈ | ≈ | ≈ R4 | ≈ R4 | ✓ |
| T4 | Regression tests | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| T5 | Dependency audit | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| T6 | Multi-device | ✓ | – | – | ≈ | – | ✓ |
| T7 | Slow networks | ✓ | ≈ | ≈ | ≈ | ≈ | ✓ |
| S1 | Auth on every endpoint | ✓ | ✓ | ≈ | – | – | ✓ (its API) |
| S2 | Secrets out of repo | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| S3 | Rotate leaked keys | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |
| S4 | Rate limiting | ✓ | ✓ | – | – | – | ✓ (its API) |
| S5 | Password reset | ✓ | ✓ if it owns auth | – | – | – | ✓ |
| S6 | Email verification | ✓ | ✓ if it owns auth | – | – | – | ✓ |
| S7 | SPF/DKIM/DMARC | ✓ if it sends email | ✓ if it sends email | ✓ if it sends email | – | – | ✓ if it sends email |
| O1 | Logging | ✓ | ✓ | ✓ | ≈ | – | ✓ |
| O2 | Error tracking | ✓ | ✓ | ✓ | ≈ | – | ✓ (crash reporting) |
| O3 | Alerting | ✓ | ✓ | ✓ | – | – | ✓ |
| O4 | Performance monitoring | ✓ | ≈ | ≈ | – | – | ✓ |
| O5 | Uptime monitoring | ✓ | ✓ | ≈ | – | – | via its API |
| O6 | Staging | ✓ | ✓ | ✓ | – | – | ≈ |
| O7 | Rollback plan | ✓ | ✓ | ✓ | ≈ R3 | ≈ R3 | ≈ |
| O8 | Database backups | ✓ if it has a DB | ✓ if it has a DB | ✓ if it has a DB | – | – | via its API |
| O9 | Restore tested | ✓ if it has a DB | ✓ if it has a DB | ✓ if it has a DB | – | – | via its API |
| O10 | Caching | ✓ | ≈ | – | – | – | ≈ |
| O11 | Spend caps | ✓ | ✓ | ✓ | – unless it calls paid APIs with your key | – | ✓ |
| A1–A4 | LLM items | ✓ if it calls a model | ✓ if it calls a model | ✓ if it calls a model | ✓ if it calls a model | ≈ | ✓ if it calls a model |
| P1 | SEO/AEO | ✓ | – | – | ≈ docs site | ≈ docs site | ≈ store listing |
| P2 | OG images + favicon | ✓ | – | – | ≈ docs site | ≈ docs site | ≈ app icon |
| P3 | Image optimization | ✓ | – | – | – | – | ✓ |
| P4 | 404/500/empty/loading | ✓ | ≈ | – | ≈ | ≈ | ✓ |
| P5 | Accessibility | ✓ | – | – | ≈ | – | ✓ |
| P6 | Analytics | ✓ | ≈ | – | ≈ opt-in only | – | ✓ |
| P7 | Custom domain + SSL | ✓ | ✓ | – | – | – | via its API |
| P8 | Privacy policy + terms | ✓ | ✓ if it has external users | – | ≈ if it collects telemetry | ≈ license instead | ✓ |
| P9 | Cookie consent | ✓ | – | – | – | – | ≈ tracking consent |
| P10 | Account + data deletion | ✓ | ✓ if it owns accounts | – | – | – | ✓ |
| P11 | README | ✓ | ✓ | ✓ | ✓ | ✓ | ✓ |

## What "≈" means per profile

**API / backend service**
- T3: smoke tests hit the health endpoint and 2–4 core endpoints with a test credential after each deploy.
- T7: becomes client-facing resilience. Timeouts on outbound calls, idempotency keys on writes so client retries are safe.
- O4: latency per endpoint and slow queries. No real-user web vitals.
- O10: HTTP cache headers on public GETs, `private`/`no-store` on anything authenticated, ETags where useful.
- P4: one consistent error format (status code, machine-readable code, human message, request ID), no stack traces, a JSON 404 for unknown routes, pagination with empty results.
- P6: API usage metrics per key or customer (calls, errors, latency), not page analytics.

**Worker / data pipeline**
- T3: a post-deploy run on a small known input with output checks.
- T7: retries with backoff on upstream calls, dead-letter queue for poison messages.
- S1: every trigger (queue, webhook, cron endpoint) authenticates its source.
- O4: job duration, throughput, queue depth and lag.
- O5: a dead-man's switch (the job pings a monitor on success, and silence alerts), not an HTTP uptime check.

**CLI tool**
- T6: supported operating systems and architectures tested in CI.
- T7: network commands have timeouts, clear offline errors, and resumable downloads where it matters.
- O1/O2: a `--verbose`/debug flag. Any crash reporting or telemetry is opt-in and disclosed.
- P4: clear error messages, non-zero exit codes, no raw stack traces unless in debug mode.
- P5: works without color (`NO_COLOR`), readable by screen readers (no meaning carried only by color or spinners).
- P8: only if it collects telemetry. Then disclose what's collected and how to opt out.

**Library / SDK**
- A1–A4: apply only if the library wraps a model; then A3/A4 are about exposing timeouts, retries and token-limit errors to callers.
- P4: typed, documented errors. Never swallow exceptions.
- P8: replaced by a clear LICENSE and a statement of any data the library sends anywhere.

**Mobile app**
- O6: internal/beta distribution track (TestFlight, internal testing) as staging.
- O7: no instant rollback of a shipped binary. The plan is server-side feature flags, a kill switch, and a forced-update path.
- O10: offline caching and cache invalidation on app update.
- P1/P2: store listing metadata and app icon set instead of SEO/OG.
- P9: platform tracking-consent rules (verify current store rules live).

## Release items (CLI tool, library/SDK)

The core 40 items don't cover publishing. For CLI and library profiles, audit these too, with the same statuses, severity and evidence rules. Owner: `test-and-ci`.

| ID | Item | PASS looks like | Severity if missing |
|---|---|---|---|
| R1 | Versioning + changelog | Semantic (or documented) versioning, a changelog entry per release, breaking changes called out | MEDIUM |
| R2 | Release pipeline | Releases are built and published from CI on a tag, not from a laptop. Registry token scoped to publish only and stored in CI secrets. 2FA on the registry account. Provenance or signing where the registry supports it (verify live). | HIGH |
| R3 | Bad-release recovery | A documented way to deprecate or yank a bad version and publish a fix quickly. A tested understanding of what the registry allows (verify live: some registries don't allow deleting versions). | HIGH |
| R4 | Install smoke test | CI installs the *published or packed* artifact into a clean environment and runs a basic command or import | HIGH |
| R5 | Package contents | The published artifact contains only what's intended: no secrets, `.env`, tests, fixtures or source maps you didn't mean to ship. Checked with the packager's dry-run/list command. | BLOCKER if secrets are found, otherwise MEDIUM |
