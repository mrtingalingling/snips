---
name: ops-reliability
description: Audits and then, with approval, sets up the operational safety net for a live app in any stack — structured logging, error tracking, alerting, performance monitoring, uptime monitoring, a staging environment, a rollback plan for bad deploys, database backups, an actually-tested restore, caching, and spend caps with billing alerts. Use when someone is about to go live or already is and asks about monitoring, observability, Sentry or similar, logs, getting paged, downtime, staging vs production, undoing a bad deploy, backups, disaster recovery, slow pages, CDN or cache headers, or surprise cloud or API bills — even if they just say "what if something breaks at 3am". Part of the ship-readiness family (items O1–O11).
---

# Ops and Reliability

Code that works today still meets bad deploys, outages, data loss and runaway bills. This skill makes sure that when one of those happens, someone finds out quickly, can see what went wrong, and can get back to a good state. Audit, report, implement only what's approved.

## Workflow

1. **Detect the stack and the providers.** Host, database and its provider, CDN, every pay-per-use service (from dependencies and env var names). If the `ship-readiness-audit` skill is installed (workspace `.agents/skills/` or global `~/.gemini/config/skills/`), use the `triage_scan.py` script in its `scripts/` folder (run it with `--help` first).
2. **Audit O1–O11.** Most of these live partly or wholly outside the repo: dashboards, provider settings, DNS. Mark what you can't see as UNKNOWN and write the exact question to close it ("In your DB provider's console, are automated backups on, and what's the retention?"). Don't infer "no backups" from "no backup code".
3. **Report and stop.** Use the ship-readiness table. Order by consequence: backups, restore test, spend caps and rollback first, because those failures are unrecoverable or unbounded. Ask what to implement. **Skip this step for items the user already approved in this conversation** (through `/ship-check` or the audit's Phase 4): go straight to step 4.
4. **Implement approved items** one per branch or commit (`reviewable-diffs` if installed). Every Human-only item gets a runbook with exact steps and a verification step. Check vendor setup steps in current docs (`live-source-verification`). Dashboards and SDK init APIs change often.
5. **Verify**: trigger a test error and see it arrive, curl headers, run the restore drill, and so on. Record the evidence.

Several items depend on others. Do error tracking (O2) and the health endpoint (O5) before alerting (O3). Set up staging (O6) before rehearsing rollback (O7) or restores (O9).

## Working rules

- **App type first.** Use the app profile from the audit if one exists, otherwise classify the project with `resources/app-profiles.md` in the `ship-readiness-audit` skill. Mark items N/A for the profile instead of reporting them MISSING, and use the profile's changed meaning where the table says so.
- **Change discipline.** Use `reviewable-diffs` if installed (workspace `.agents/skills/` or global `~/.gemini/config/skills/`). Otherwise: one approved item per branch or commit; change only the lines that item needs; no formatting, renaming or refactoring churn; no features nobody asked for; add a test that fails without the fix and passes with it; list each file touched and why.
- **Current facts.** Use `live-source-verification` if installed. Either way, apply the citation rule: any vendor-specific or time-sensitive fact (pricing, plan limits, config keys, CLI flags, SDK and error names, DNS values, legal or store rules) carries an inline [Title](URL) retrieved this session. No URLs from memory. If a source is unreachable, write "Source unavailable" and tag the claim [Unknown].
- **Confidence tags.** Tag each finding's evidence with exactly one of [Cited], [Certain], [Likely], [Guessing], [Unknown] ([Cited]: checked against a live source this session, with the link; [Certain]: seen directly this session; [Likely]: strong inference; [Guessing]: hypothesis only; [Unknown]: name the gap and what would close it). PASS requires [Certain] or [Cited].
- **Irreversible steps need their own yes**, even for approved items: rewriting git history, rotating or revoking live keys, DNS changes, production deploys, restoring or deleting data, billing changes.

## Item playbooks

Detailed guidance for each item is in `resources/items.md`. Read only the sections for the items in scope, plus the **Current requirements** section at the end, which holds externally verified standards and dates:

- O1 — Logging
- O2 — Error tracking
- O3 — Alerting
- O4 — Performance monitoring
- O5 — Uptime monitoring
- O6 — Staging environment
- O7 — Rollback plan for bad deploys
- O8 — Database backups
- O9 — Actually test a restore
- O10 — Caching
- O11 — Spend caps and billing alerts
