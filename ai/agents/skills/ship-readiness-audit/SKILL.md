---
name: ship-readiness-audit
description: Audits an app or repo for production readiness across 40 launch items — CI/CD, tests, security, auth, secrets, monitoring, backups, rollback, spend caps, SEO, accessibility, legal pages, LLM hardening and more — produces an evidence-backed gap report ranked by severity, then implements only the items the user approves. Use when someone asks if their app is ready to launch, ship, go live or go to production, wants a pre-launch or go-live checklist, asks "what am I missing", wants to harden a vibe-coded or prototype app, or is moving from side project to real users — even if they don't say "audit". Hands deep work on one area to test-and-ci, security-hardening, ops-reliability, llm-hardening or launch-polish.
---

# Ship-Readiness Audit

The point of this skill is to turn "is this ready?" into a short, ranked, evidence-backed list of what is actually missing, and then fix only what the user approves. A prototype usually works on the happy path; what's missing is what happens when things go wrong (bad deploy, leaked key, DB loss, model outage, surprise bill) and the unglamorous launch work (legal pages, SEO, accessibility). This skill finds those gaps fast and routes depth to five sibling skills.

## Sibling skills and how work flows

| Area | Skill | Item IDs |
|------|-------|----------|
| Testing and CI | `test-and-ci` | T1–T7 |
| Security and auth | `security-hardening` | S1–S7 |
| Operations and reliability | `ops-reliability` | O1–O11 |
| LLM features | `llm-hardening` | A1–A4 |
| Launch polish and compliance | `launch-polish` | P1–P11 |

The full 40-item checklist, with what counts as evidence for each, is in `resources/checklist.md`. Read it before auditing.

Three other skills may be installed in the workspace (`.agents/skills/`) or globally (`~/.gemini/config/skills/`). This one leans on them rather than duplicating them:

- **reviewable-diffs** governs *how* approved changes land: branch per change, truth table, tests first, minimal diff. The approved item list from this audit becomes the requirement list you hand it. Because the user explicitly approved those items, adding logging or error handling for an approved item is in scope, not scope creep.
- **live-source-verification** governs any fact that can go stale: current CLI flags, library versions, provider DNS records, privacy/cookie law. Don't write vendor config from memory; check the vendor's current docs.

- **adversarial-advisor** reviews a decision to ship with open BLOCKER or HIGH items (see "Shipping with open items").

If they aren't installed, the plugin still works on its own:

- If `reviewable-diffs` isn't installed, follow these rules directly: one approved item per branch or commit; change only the lines that item needs; no formatting, renaming or refactoring churn; no features nobody asked for; add a test that fails without the fix and passes with it; list each file touched and why.
- If `live-source-verification` isn't installed, follow the citation rule directly: break the claim into its individual facts, check each against the primary source, cite it inline, and stop searching once every fact the decision depends on has a source.
- The "Shipping with open items" section below spells out the adversarial review.

## Citation rule

Any vendor-specific or time-sensitive fact in a report, runbook, config or code comment carries an inline [Title](URL) retrieved in this session. That includes pricing, plan limits, whether a hard spend cap exists, DNS record values, config keys, CLI flags, SDK method or error names, model identifiers, store rules, and privacy or accessibility requirements. Prefer the vendor's or standard body's own docs over blogs and aggregators.

Never write a URL from memory. If the source can't be reached, write "Source unavailable", tag the claim [Unknown], and don't substitute a plausible-looking alternative. Where sources conflict, show both with attribution instead of picking one silently.

## Confidence tags

Status says whether an item is done. A confidence tag says how sure the finding is. Tag every finding's evidence with exactly one:

- **[Cited]**: verified against a live source in this session; carries an inline [Title](URL). The link is the tag.
- **[Certain]**: seen directly in this session: file and line, command output, a response header.
- **[Likely]**: strong inference, could be wrong at the edges (e.g. "vulnerable function appears unreachable from request handlers").
- **[Guessing]**: a hypothesis, not an input. Never the basis for a status other than UNKNOWN.
- **[Unknown]**: couldn't find out. Name the gap and what would close it.

Rules: PASS requires [Certain] or [Cited] evidence. A PARTIAL or MISSING built on [Likely] says so. UNKNOWN always pairs with [Unknown]. If a section of the report is mostly [Likely] or [Guessing] (common in description mode), say so in its first line.

## Phase 0 — Establish access

What you can verify depends on what the agent can reach. Say which mode you're in at the top of the report.

- **Repo mode** (the normal case in Antigravity: an open workspace with a terminal): you can read files, run scripts, run the test suite. If browser use is available, you can also load deployed pages to check rendered output, headers and layout.
- **Description mode** (no workspace or no terminal access): you can only reason from what the user tells you. Ask once for a zip, a repo link, or the key files (package manifest, CI config, route files, env example). If they decline, proceed and mark repo-dependent items UNKNOWN.

Many items live outside the repo: uptime monitors, billing alerts, backups, DNS records, provider dashboards. You can't see those from code. Absence in the repo is not evidence of absence. Mark them UNKNOWN and give the user the one question or command that would settle it.

### App type

Classify the project with `resources/app-profiles.md` (web app, API/backend, worker/pipeline, CLI, library/SDK, mobile, or a combination). Record the profile and the evidence for it at the top of the report. The profile decides which of the 40 items apply, which change meaning, and whether the release items R1–R5 are added. If the signals conflict, ask the user one question rather than guessing.

## Phase 1 — Triage scan

In repo mode, run the bundled scanner first. It's deterministic and saves you rediscovering the stack every time. Treat it as a black box: run it with `--help` first rather than reading the source.

```bash
python3 scripts/triage_scan.py --help
python3 scripts/triage_scan.py <repo-root> > triage.json
```

The path is relative to this skill's folder, usually `.agents/skills/ship-readiness-audit/` in the workspace or `~/.gemini/config/skills/ship-readiness-audit/` globally.

It reports signals, not verdicts: detected languages, frameworks, hosting and CI config, test frameworks, observability SDKs, LLM SDKs, presence of robots.txt, sitemap, favicon, README, tracked `.env` files, and redacted secret-pattern hits in the working tree. A signal like "Sentry SDK in dependencies" means "look here", not "error tracking PASS". Confirm each with the evidence the checklist asks for.

The scanner only sees the current working tree. Secrets in git history need a history scanner (see `security-hardening`, S2/S3).

## Phase 2 — Audit every item

Walk all 40 items in `resources/checklist.md`, plus R1–R5 from `resources/app-profiles.md` for CLI and library profiles. Items the profile marks N/A are recorded as N/A with the profile as the reason. Items it marks as changed are audited against the changed meaning. For each, record exactly one status:

- **PASS**: you saw evidence it's done and correct. Cite it (file path and line, command output, or config key).
- **PARTIAL**: exists but incomplete or wrong in a way that matters. Say what's missing.
- **MISSING**: you looked where it would be and it isn't there.
- **UNKNOWN**: you couldn't look (external dashboard, no repo access, needs a human). Name the question or command that would close it.
- **N/A**: doesn't apply. Say why in a few words (e.g. "no LLM calls", "sends no email", "no non-essential cookies").

The reason for this discipline: a PASS without evidence is how launch checklists lie. A reviewer should be able to check every PASS in under a minute.

Assign severity by consequence, not effort:

- **BLOCKER**: can cause unrecoverable data loss, a security breach, legal exposure, or an unbounded bill. Examples: secret in the repo, endpoint missing auth, no backups, no spend cap on a pay-per-use API exposed to the public.
- **HIGH**: users will hit it or you'll be blind when they do. No error tracking, no rollback path, broken password reset.
- **MEDIUM**: degrades quality or growth. No sitemap, no OG images, unoptimized images.
- **LOW**: nice to have at launch.

Also estimate fix size (S: under an hour, M: a few hours, L: a day or more) and who does it. **Agent** means it's a code or config change in the repo. **Human** means it needs a dashboard, DNS, billing, a real device, or legal review. Many items are mixed. Split them so the human part is an explicit task, not a silent gap.

## Phase 3 — Report, then stop

Use this structure exactly, so reports are comparable across runs and repos:

```markdown
# Ship-readiness audit — <project> — <YYYY-MM-DD>
Mode: repo | description   ·   App type: <profile(s), with evidence>   ·   Stack: <detected, with evidence>

## Summary
<2–4 sentences: overall state, count by severity, the single most important fix>

## Blockers
| ID | Item | Status | Evidence | Confidence | Fix | Size | Who |

## High
(same table)

## Medium / Low
(same table, may be collapsed)

## Unknowns to resolve
<one line each: the question or command that would settle it>

## Proposed order
<numbered: blockers first, then items that unblock others (e.g. staging before smoke tests, error tracking before alerting)>

## Not applicable
<ID — reason>
```

Offer to save the report to `docs/ship-readiness.md` so later re-audits have a baseline. Then stop and ask which items to implement. Accept answers like "all blockers", "S1, O8, O9", or "everything the agent can do". Don't start changing code before approval. An audit the user didn't ask to act on should stay an audit, and some fixes (key rotation, DNS, deleting data) have consequences outside the repo.

## Shipping with open items

If the user proposes launching with any open BLOCKER or HIGH item, or asks "should we ship?", "is this good enough?" or similar, hand the report to `adversarial-advisor` if installed. Otherwise review it directly:

1. Lead with the single most dangerous open item and its worst realistic consequence.
2. Name what accepting it trades away, and what would have to be true for that trade to be reasonable (a mitigation in place, a time limit, a small user base, no payments yet).
3. No reassurance or validation language. If the plan is sound, confirm that in one line instead of inventing objections.
4. The decision is the user's. Record it in the report as an accepted risk with an owner and a revisit date.

## Phase 4 — Implement approved items

- Group approved items by sibling skill. If the sibling skill is installed, read its SKILL.md and the matching section of its `resources/items.md`, and follow that guidance for those items. The approval given here covers them: skip the sibling's own report-and-stop step for those items. If it isn't installed, use the "Fix" notes in `resources/checklist.md`.
- One logical change per branch or commit (reviewable-diffs). Dependency bumps, formatter config, and feature work never share a commit.
- For every Human task, write a short runbook (exact steps, what "done" looks like, how to verify) instead of just saying "configure X".
- Verify each fix the way the checklist's evidence column describes. Run the test, curl the endpoint, render the page. Re-mark the item with the new evidence.
- Finish with the same report table, statuses updated, plus a list of files touched and remaining Human tasks.

## Re-audits

When the user comes back later, re-run Phase 1 and 2 and show a diff of statuses against the previous report if one is in the conversation or repo (e.g. `docs/ship-readiness.md`). Offer to save the report there so the next audit has a baseline.
