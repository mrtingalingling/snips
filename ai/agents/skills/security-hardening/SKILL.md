---
name: security-hardening
description: Audits and then, with approval, fixes the security basics of a web app in any stack — auth and authorization on every endpoint, secrets kept out of the repo, rotating leaked keys, rate limiting, password reset and email verification flows, and SPF/DKIM/DMARC email authentication. Use when someone asks whether their app is secure, wants to check auth or permissions, worries a user could see another user's data, committed an API key or .env file, needs to rotate credentials, is getting abused by bots or brute-force, has broken reset or verification emails, or has emails landing in spam — even if they don't say "security". Part of the ship-readiness family (items S1–S7).
---

# Security Hardening

This covers the security failures that most often sink young apps: an endpoint nobody remembered to protect, a key in the repo, no brake on brute-force, and account flows that leak or break. It's defensive work on the user's own app. Audit first, report, then fix only what's approved.

## Workflow

1. **Detect the stack and the auth model.** Framework, auth library or provider, session vs token, database and any row-level security, email provider. If `ship-readiness-audit` is installed (workspace `.agents/skills/` or global `~/.gemini/config/skills/`), its `scripts/triage_scan.py` (run it with `--help` first) finds frameworks and redacted secret-pattern hits in the working tree.
2. **Audit S1–S7** with PASS / PARTIAL / MISSING / UNKNOWN / N/A and evidence.
3. **Report and stop.** Use the ship-readiness table, with security blockers at the top. Ask which items to implement.
4. **Implement approved items** one per branch or commit (`reviewable-diffs` if installed). Look up current provider settings, DNS records and library APIs in official docs (`live-source-verification`).
5. **Verify** with tests or commands that would fail if the fix were missing, and record them as evidence.

Never print a full secret in a report, log, commit message or chat. Show a short prefix and the length, like the triage scanner does. A report that repeats the leaked key is a second leak.

## S1 — Auth on every endpoint

The failure mode is rarely "no auth at all". It's one forgotten route, a server action that skips the middleware, or an endpoint that checks *who you are* but not *whether this record is yours*. The last one is an insecure direct object reference (IDOR), and it's the most common serious bug in apps built fast.

Build a **route matrix** before judging anything. Enumerate every entrypoint: page routes, API handlers, server actions/functions, RPC or tRPC procedures, GraphQL resolvers (each field that resolves data), webhooks, cron or job endpoints, file/upload URLs, and admin routes. Discover them from the framework's routing conventions and from grep, not from memory of what the app "should" have.

| Route | Method | Intended access | Authn enforced at | Authz/ownership check | Test |
|-------|--------|-----------------|-------------------|-----------------------|------|

Intended access is one of: public, any signed-in user, owner of resource, role (admin…), or machine (webhook/cron with signature or secret).

What to look for:
- **Default-deny.** Is auth enforced centrally (middleware or a wrapper all handlers use) with an explicit public allowlist? Or does each handler opt in? Opt-in is where routes get forgotten. Recommend default-deny, but don't refactor all routing without approval. It's a large change.
- **Ownership.** For every route that takes an ID, confirm the query is scoped to the current user or tenant (e.g. `WHERE id = ? AND owner_id = ?`), or that a policy layer (row-level security, policy objects) enforces it. Check that RLS is actually enabled on each table if the platform uses it. Disabled RLS with a public client key is a full data exposure.
- **Client-trusted fields.** User ID, role, price or tenant taken from the request body instead of the session.
- **Webhooks** verify the provider's signature with the raw body and reject replays where the provider supports it.
- **Cron/job endpoints** require a secret or platform-level protection.

Tests: for each non-public route, assert that unauthenticated gets 401 (or redirect to login for pages), and that an authenticated *different* user gets 403 or 404 on someone else's resource. Two users and one resource are enough to catch IDOR.

## S2 — Secrets out of the repo

Check all of these:
- Tracked env files: `git ls-files | grep -iE '(^|/)\.env'`. Example/template files with placeholders are fine.
- Working-tree patterns: the triage scanner, or grep for known key prefixes and `PRIVATE KEY` blocks.
- **Git history**: a secret deleted in a later commit is still in the history and still leaked if the repo was ever pushed anywhere shared. Use a history-aware scanner such as gitleaks or trufflehog if available, and check its current flags. If you can't run one, mark history UNKNOWN and give the command.
- **Client bundles**: frameworks expose env vars with certain prefixes to the browser (public-prefixed vars). Any secret under a public prefix ships to every visitor. Also check for keys hardcoded in front-end code and in mobile app bundles.
- CI logs and config that echo env vars.

Fix: move secrets to env vars from the host's or CI's secret store. Add ignore rules. Commit an `.env.example` with names and placeholders only. Add a pre-commit or CI secret scan so it doesn't recur. Rewriting history to purge a secret is optional and disruptive (force-push, every clone affected). It does **not** make the secret safe again, so rotation (S3) is required either way.

## S3 — Rotate leaked keys

Treat any secret that was ever committed to a pushed repo, pasted in a public place, shipped in a client bundle, or printed in shared logs as compromised. Rotation happens at the provider's dashboard, so it's a Human task. Your job is a precise runbook per key:

1. Create a new key at the provider, with the narrowest scope that works.
2. Update every place it's stored: host env vars for each environment, CI secrets, local `.env` files of teammates, any other services that call with it.
3. Redeploy so running instances pick it up.
4. Revoke the old key. Do this after the deploy, or the app breaks. Do it promptly, or the rotation was pointless.
5. Check the provider's usage or audit logs for activity you don't recognize between leak date and revocation, and report anything found.

Order the runbook by blast radius: cloud root/admin credentials, payment provider secrets, database URLs, LLM/API keys with billing, then everything else. Mark each key PASS only when the user confirms the old one is revoked.

## S4 — Rate limiting

Protect at minimum: login, signup, password reset request, verification-email resend, any endpoint that sends email or SMS, and any endpoint that costs money per call (LLM, image generation, third-party paid APIs). Also consider upload and search endpoints.

Design points:
- Limit by IP *and* by account or identifier (e.g. per email for reset requests). IP-only limits are evaded by botnets and also punish shared NATs.
- **Use a shared store** (Redis or a hosted limiter, or the platform's edge rate limiting). In-memory limiters reset per instance and don't work on serverless or multi-instance deploys.
- Return 429 with a `Retry-After` header. Show a human message in the UI.
- Login: slow down or require extra verification after repeated failures. Don't lock accounts permanently, since that becomes a denial-of-service tool against your users.

Tests: exceed the limit in a test and assert 429.

## S5 — Password reset works (and is safe)

Check the flow end to end:
- The request endpoint returns the same response whether or not the email exists. It's rate-limited (S4).
- The token is random and long, single-use, expires (minutes to an hour, not days), and is stored hashed, not plaintext.
- The reset link is built from a **configured base URL**, never the request's Host header. Host-header injection can send users a link to an attacker's domain.
- Completing the reset invalidates other sessions and outstanding reset tokens. It enforces the same password rules as signup, and sends a "your password was changed" notice.
- If the app uses a hosted auth provider, check its reset settings (redirect URL allowlist, token expiry) rather than reimplementing.

"Works" needs a real run: request a reset in staging, receive the email in a real inbox, follow the link, set the password, log in with it, and confirm the old session is gone. That's a Human or staging task if you can't receive email. Write it as a checklist. Also check the email isn't landing in spam (S7).

## S6 — Email verification works

- Decide what unverified users can do, and enforce it on the server, not just by hiding UI.
- Tokens expire and are single-use. Resend is rate-limited.
- Changing email requires re-verification before the new address is used for login or reset. Notify the old address.
- The verification link uses the configured base URL (same reason as S5).
- Run it end to end in staging, same as S5.

## S7 — SPF, DKIM, DMARC

N/A if the app sends no email from its own domain. Otherwise these are DNS records, so the edits are a Human task. You audit and write exact records.

Audit with `dig +short TXT <domain>`, `dig +short TXT _dmarc.<domain>`, and the DKIM selector record each provider specifies (`dig +short TXT <selector>._domainkey.<domain>`). Without a terminal, give the user these commands or an online checker.

- **SPF**: exactly one `v=spf1` TXT record on the sending domain, including every service that sends as that domain (app email provider, workspace mail, marketing tool). Multiple SPF records is an error. There's a DNS-lookup limit, so check it when many includes stack up.
- **DKIM**: enabled for each sending provider, with the exact records from that provider's current docs. Never guess selector names or values.
- **DMARC**: a `_dmarc` TXT record. Start with `v=DMARC1; p=none; rua=mailto:<reports-address>` to collect reports without affecting delivery. Once reports show all legitimate mail passes, move to `quarantine` then `reject`. Tightening too early silently drops real mail, so call this out as a staged plan, not a one-shot change.

Verify by sending a test email to a mailbox that shows authentication results (most webmail "show original" views do) and recording SPF/DKIM/DMARC pass or fail.
