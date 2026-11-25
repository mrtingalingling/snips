---
name: ship-check
description: >
  This skill should be used when the user types "/ship-check" or "ship check", or asks to "run a ship check",
  "get this production-ready", "go through the launch checklist and fix things", or "audit and fix
  before launch". Entry point of the ship-readiness family: runs the audit, gets approval once, then
  routes each approved item to the skill that owns it. Accepts an optional argument: "audit-only",
  a severity ("blockers", "high"), an area ("tests", "security", "ops", "ai", "polish", "release"),
  or one or more item IDs or names (e.g. "S4", "O8 O9", "rate limiting").
metadata:
  version: "0.4.0"
---

# /ship-check

Orchestrate the ship-readiness family. Don't implement anything in this skill itself. Route each item to the skill that owns it.

| Area | Skill | Item IDs | `/ship-check` area argument |
|---|---|---|---|
| Testing and CI | `test-and-ci` | T1–T7 | `tests` |
| Security and auth | `security-hardening` | S1–S7 | `security` |
| Operations and reliability | `ops-reliability` | O1–O11 | `ops` |
| LLM features | `llm-hardening` | A1–A4 | `ai` |
| Release (CLI and library only) | `test-and-ci` | R1–R5 | `release` |
| Launch polish and compliance | `launch-polish` | P1–P11 | `polish` |

## Arguments

| Argument | Behavior |
|---|---|
| (none) | Full audit, then ask which items to fix |
| `audit-only` | Full audit, report, stop |
| `blockers` / `high` | Full audit, then propose only items at that severity or above |
| an area (`tests`, `security`, `ops`, `ai`, `polish`, `release`) | Audit only that area's items, then ask |
| item IDs or names (`S4`, `O8 O9`, `rate limiting`) | Skip the full audit. Assess just those items with evidence, then ask |

## Flow

1. **Audit.** Load `ship-readiness-audit` and follow its Phases 0–3, scoped by the argument. The app type it records decides which items are in play: never propose fixing an item the profile marks N/A. If a previous report exists (`docs/ship-readiness.md` or earlier in the conversation), show the status diff.
2. **Approve once.** Present the proposed order from the report, split into:
   - *Agent can do*: code or config changes in the repo
   - *You need to do*: dashboards, DNS, billing, key revocation, real devices, legal review
   Ask which to implement, offering up to four groupings as choices (for example "All blockers the agent can do", "All blockers + high", "Security items only"). Free-text answers like "S1, O8, O9" also work.
3. **Fix, one item at a time.** For each approved item, in the proposed order:
   - Load the owning skill's SKILL.md, then read only that item's section of its `resources/items.md`.
   - Tell the owning skill the item is pre-approved, so it skips its report-and-stop step.
   - Implement it as its own change set (`reviewable-diffs` if installed, otherwise the fallback rules in the audit skill) and verify it the way the checklist's evidence column says.
   - For *You need to do* items, write the runbook (exact steps, what done looks like, how to verify) instead of a vague "configure X".
4. **Re-mark.** Update each item's status and evidence in the report as you go.
5. **Ship decision.** If the user wants to launch with any BLOCKER or HIGH item still open, or asks whether it's good enough to ship, follow the audit's "Shipping with open items" section (it hands off to `adversarial-advisor` when installed). Record the outcome as an accepted risk with an owner and revisit date.
6. **Close out.** Give the new severity counts, the files touched, and the remaining *You need to do* checklist. Offer to save or update `docs/ship-readiness.md`.

## Rules

- Approval in step 2 covers the code change for those items only. Anything irreversible or shared still needs its own explicit yes at the moment it happens: rewriting git history, rotating or revoking live keys, DNS changes, production deploys, restoring or deleting data, or billing changes.
- Never skip the audit except for the item-ID argument. Fixing without a baseline makes "done" unverifiable.
- If nobody is there to answer (scheduled or unattended run), run the audit and only non-destructive code fixes for BLOCKER and HIGH items, then list everything else as pending decisions at the top of the report.
