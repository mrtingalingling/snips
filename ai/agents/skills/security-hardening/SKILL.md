---
name: security-hardening
description: Audits and then, with approval, fixes the security basics of a web app in any stack — auth and authorization on every endpoint, secrets kept out of the repo, rotating leaked keys, rate limiting, password reset and email verification flows, and SPF/DKIM/DMARC email authentication. Use when someone asks whether their app is secure, wants to check auth or permissions, worries a user could see another user's data, committed an API key or .env file, needs to rotate credentials, is getting abused by bots or brute-force, has broken reset or verification emails, or has emails landing in spam — even if they don't say "security". Part of the ship-readiness family (items S1–S7).
---

# Security Hardening

This covers the security failures that most often sink young apps: an endpoint nobody remembered to protect, a key in the repo, no brake on brute-force, and account flows that leak or break. It's defensive work on the user's own app. Audit first, report, then fix only what's approved.

## Workflow

1. **Detect the stack and the auth model.** Framework, auth library or provider, session vs token, database and any row-level security, email provider. If the `ship-readiness-audit` skill is installed (workspace `.agents/skills/` or global `~/.gemini/config/skills/`), the `triage_scan.py` script in its `scripts/` folder (run it with `--help` first) finds frameworks and redacted secret-pattern hits in the working tree.
2. **Audit S1–S7** with PASS / PARTIAL / MISSING / UNKNOWN / N/A and evidence.
3. **Report and stop.** Use the ship-readiness table, with security blockers at the top. Ask which items to implement. **Skip this step for items the user already approved in this conversation** (through `/ship-check` or the audit's Phase 4): go straight to step 4.
4. **Implement approved items** one per branch or commit (`reviewable-diffs` if installed). Look up current provider settings, DNS records and library APIs in official docs (`live-source-verification`).
5. **Verify** with tests or commands that would fail if the fix were missing, and record them as evidence.

Never print a full secret in a report, log, commit message or chat. Show a short prefix and the length, like the triage scanner does. A report that repeats the leaked key is a second leak.

## Working rules

- **App type first.** Use the app profile from the audit if one exists, otherwise classify the project with `resources/app-profiles.md` in the `ship-readiness-audit` skill. Mark items N/A for the profile instead of reporting them MISSING, and use the profile's changed meaning where the table says so.
- **Change discipline.** Use `reviewable-diffs` if installed (workspace `.agents/skills/` or global `~/.gemini/config/skills/`). Otherwise: one approved item per branch or commit; change only the lines that item needs; no formatting, renaming or refactoring churn; no features nobody asked for; add a test that fails without the fix and passes with it; list each file touched and why.
- **Current facts.** Use `live-source-verification` if installed. Either way, apply the citation rule: any vendor-specific or time-sensitive fact (pricing, plan limits, config keys, CLI flags, SDK and error names, DNS values, legal or store rules) carries an inline [Title](URL) retrieved this session. No URLs from memory. If a source is unreachable, write "Source unavailable" and tag the claim [Unknown].
- **Confidence tags.** Tag each finding's evidence with exactly one of [Cited], [Certain], [Likely], [Guessing], [Unknown] ([Cited]: checked against a live source this session, with the link; [Certain]: seen directly this session; [Likely]: strong inference; [Guessing]: hypothesis only; [Unknown]: name the gap and what would close it). PASS requires [Certain] or [Cited].
- **Irreversible steps need their own yes**, even for approved items: rewriting git history, rotating or revoking live keys, DNS changes, production deploys, restoring or deleting data, billing changes.

## Item playbooks

Detailed guidance for each item is in `resources/items.md`. Read only the sections for the items in scope, plus the **Current requirements** section at the end, which holds externally verified standards and dates:

- S1 — Auth on every endpoint
- S2 — Secrets out of the repo
- S3 — Rotate leaked keys
- S4 — Rate limiting
- S5 — Password reset works (and is safe)
- S6 — Email verification works
- S7 — SPF, DKIM, DMARC
