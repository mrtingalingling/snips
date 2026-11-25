# Ops and Reliability: item playbooks

## O1 — Logging

Good logs answer "what happened to this request?" in one search. Check for:
- Structured output (JSON or key=value) with timestamp, level, message, and a **request/correlation ID** carried through the request and into downstream calls.
- Meaningful events logged at boundaries: incoming request summary, outbound calls with duration and status, auth failures, job start/finish.
- **No secrets or personal data** in log payloads: tokens, passwords, full request bodies, auth headers, raw emails/phone numbers where avoidable. Add a redaction step.
- Logs shipped somewhere searchable with retention the user knows about. Host-provided logs often keep only a short window, so check and report it.

Prefer the stack's idiomatic structured logger. Don't replace an existing logger without approval.

## O2 — Error tracking

An error tracker turns "a user said it's broken" into a stack trace with context. Check:
- SDK initialized on the server and the client (front-end errors are invisible otherwise), plus background workers.
- DSN/key from env, environment tag (production/staging), and **release/version tag** set from CI (commit SHA). Without release tags you can't tell whether a deploy caused a spike.
- Source maps uploaded for minified front-end code, and not publicly served if the user cares about that.
- PII scrubbing settings reviewed. User IDs are usually fine. Passwords and tokens never.
- Unhandled promise rejections and framework error boundaries captured.

Verify by throwing a deliberate test error in staging and confirming it appears with the right release and environment.

## O3 — Alerting

Alerts should be few, actionable, and owned. Propose a starter set and let the user choose thresholds:
- Site down (from O5)
- Error rate or new-issue spike (from O2)
- Background jobs failing or not running (dead-man's switch, where the job pings a monitor on success)
- Spend threshold crossed (from O11)
- Optionally: p95 latency regression (from O4)

For each: the signal, the threshold, where it goes (email, chat channel, phone), and who responds. Noisy alerts get muted and then miss the real incident. If an alert fires and nobody would act, drop it. Test-fire every alert once (a forced error, a stopped health check, a low temporary spend threshold) and have its owner confirm it arrived. An alert nobody has seen arrive isn't done. Channel configuration is a Human task. Alert rules can live in code if the provider supports infrastructure-as-code.

## O4 — Performance monitoring

Two views:
- **Server**: request latency (p50/p95), throughput, slow endpoints, slow DB queries. Usually via the error tracker's tracing, an APM, or OpenTelemetry. Sample traces in production rather than tracing 100% unless volume is tiny.
- **Real users (web)**: Core Web Vitals and page-load timing from real browsers, via the host's analytics, the error tracker, or a web-vitals reporter to an endpoint.

Report a baseline if you can measure one (a Lighthouse or equivalent run on key pages). Lab numbers aren't real-user numbers, so label them as lab.

## O5 — Uptime monitoring

Add a health endpoint (e.g. `/health` or `/api/health`) that is cheap, unauthenticated, uncached, and returns non-200 when a critical dependency (database) is unreachable. It shouldn't leak versions or internals publicly.

The monitor itself must run **outside** the hosting provider. A host's own status page doesn't tell you your app is down. Check the health endpoint and the home page, at an interval suited to the app, alerting a human (O3). Setting up the monitor account is a Human task. Write the exact URLs, expected status, and interval.

## O6 — Staging environment

Staging is where you find out a change breaks things before users do. Good looks like:
- A separate deployment (a long-lived staging branch, or per-PR previews) with **its own database** and its own secrets.
- Third-party services in test or sandbox mode: payments in test mode, email to a sink or allowlist, LLM calls with a low spend cap.
- Production-like config: same runtime versions, same build, and environment variables differing only where they must.
- Not indexed by search engines (noindex header or robots disallow; coordinate with `launch-polish` P1), and ideally behind auth.
- Never pointing at the production database. Check env var values or names per environment for this specifically. It's the classic staging disaster.

Seed data: synthetic or anonymized, never a raw copy of production personal data unless the user has a lawful reason and has thought about it.

## O7 — Rollback plan for bad deploys

The plan should let someone stressed, at night, return to the last good release in one step. Check and write `ROLLBACK.md`:
- **How**: the host's "redeploy previous" / instant rollback feature, re-running the deploy job at the previous commit, or redeploying a previous immutable image tag. Write exact commands or clicks.
- **Database migrations** are the hard part. Code can roll back. Destructive migrations can't. Recommend expand/contract: add new columns/tables first, deploy code that handles both shapes, migrate data, then remove old structures in a later release. Flag any existing migration that drops or renames in one step.
- **Feature flags** for risky features, so "turn it off" doesn't require a deploy.
- **Who decides** and how users are told (status page or message).

PASS requires a rehearsal: perform one rollback in staging, record how long it took, and note anything that surprised you.

## O8 — Database backups

Usually a provider setting, so audit by asking or by reading infrastructure config. Good looks like: automated daily (or more frequent) backups, point-in-time recovery if the provider offers it, retention the user has chosen deliberately, stored in a separate location or account from the primary, and encrypted. Also cover uploaded files (object storage) and anything else that can't be regenerated. The database isn't the only data.

If the provider doesn't offer backups on the current plan, propose a scheduled logical dump to separate storage. Treat the job and its storage credentials as production infrastructure: least privilege, failure alerting (O3).

## O9 — Actually test a restore

A backup that has never been restored is a hope, not a backup. The drill:
1. Restore the latest backup (or a point in time) into a **new, scratch** database. Never restore over production as a test.
2. Verify with checks written in advance: row counts for key tables within the expected range, latest record timestamp close to the backup time, the app (staging) boots and can read and log in against it, and a few known records exist.
3. Time it. That's the real recovery time. Compare the backup timestamp to "now", which is the real data-loss window.
4. Write `RESTORE.md`: exact steps, who has access, credentials location (not the credentials), measured times, date of last drill. Delete the scratch database afterwards.
5. Put a recurring reminder on the calendar to repeat the drill. Quarterly is a common default, and the interval is a preference, not a rule.

You can script steps 2 and 4 in repo mode. Step 1 usually needs provider access, so it's a Human task with your runbook.

## O10 — Caching

Check response headers on real URLs (`curl -sI <url>`) and the framework or CDN config:
- **Static assets with content hashes in filenames**: long-lived, immutable caching.
- **HTML**: short or revalidated caching, so deploys show up.
- **Authenticated or user-specific responses** must not be publicly cacheable (`private` / `no-store`, correct `Vary`). A CDN serving one user's page to another is a security incident, not a performance bug. Treat any sign of it as a BLOCKER.
- **API responses**: deliberate. Cache only public, idempotent GETs.
- **Application caching** (memoized DB queries, Redis) only where there's a measured slow path. Every cache needs an invalidation story.

## O11 — Spend caps and billing alerts

List every pay-per-use provider from dependencies and env var names: cloud/host, database, LLM APIs, email, SMS, maps, storage/bandwidth, search, and anything else metered. For each, write a Human runbook entry:
- Budget alert at a threshold the user picks, plus one at a higher threshold.
- A **hard cap** where the provider offers one, and the user's decision on what happens at the cap (service stops vs. keeps billing). Some providers only alert, so say which.
- Who receives the alert.

In-app (the agent can implement): per-user and global quotas on expensive endpoints (coordinate with `security-hardening` S4 and `llm-hardening`), so one abusive user or a bug in a loop can't spend the monthly budget in an hour. Also check for unbounded retry loops or recursive calls around paid APIs.

## Current requirements (verified November 2026)

These change. Before relying on one in a report or runbook, re-check the linked source and cite it.

- **O4 Web vitals targets.** Core Web Vitals are LCP, INP and CLS, judged at the 75th percentile of real visits ([web.dev](https://web.dev/articles/vitals/)). Good is LCP ≤ 2.5 s, INP ≤ 200 ms (INP replaced FID) and CLS ≤ 0.1 ([Google Search Central](https://developers.google.com/search/docs/appearance/core-web-vitals)). Report field data, not just lab scores.
- **O3/O5 Certificate expiry.** Maximum public TLS certificate lifetimes dropped to 200 days on March 15, 2026, and fall to 100 days in March 2027 and 47 days in March 2029 ([CA/B Forum SC-081v3](https://cabforum.org/2025/04/11/ballot-sc081v3-introduce-schedule-of-reducing-validity-and-data-reuse-periods/); [DigiCert](https://www.digicert.com/blog/1-year-tls-ssl-certificates-are-here-what-now)). Add a certificate-expiry check to uptime monitoring, and alert well before expiry in case auto-renewal silently fails.
