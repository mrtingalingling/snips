---
name: llm-hardening
description: Audits and then, with approval, hardens an app's AI/LLM features for production, in any stack and with any model provider — writes evals, tests for prompt injection (direct and indirect), adds fallbacks for when the model API is down, slow or rate-limited, and handles token limits and context overflow. Use when an app calls OpenAI, Anthropic, Gemini, a local model or any LLM API and someone asks about evals, testing prompts, measuring AI quality, jailbreaks or prompt injection, RAG or agent safety, retries and timeouts on model calls, 429 or overloaded errors, provider outages, context length errors, long chat histories, or truncated responses — even if they just say "make the AI part production-ready". Part of the ship-readiness family (items A1–A4).
---

# LLM Hardening

Model calls fail differently from ordinary code. Outputs vary, quality drifts when prompts or models change, any text the model reads can try to steer it, providers have outages and rate limits, and inputs grow until they don't fit. This skill audits those four risks and fixes what the user approves.

If the app makes no model calls, mark A1–A4 N/A and stop.

## Workflow

1. **Map every model call.** For each call site, record: provider and model, what goes into the prompt (system, user input, retrieved docs, tool results, files), whether tools/functions are enabled and what they can do, how output is used (shown as text, rendered as HTML/markdown, parsed as JSON, executed, used to call other systems), and the error handling around it. This map is the basis for all four items. Put it at the top of the report.
2. **Audit A1–A4** with PASS / PARTIAL / MISSING / UNKNOWN and evidence.
3. **Report and stop.** Use the ship-readiness table. Ask what to implement. **Skip this step for items the user already approved in this conversation** (through `/ship-check` or the audit's Phase 4): go straight to step 4.
4. **Implement approved items** one per branch or commit (`reviewable-diffs` if installed). SDK method names, error types, stop-reason values, token-counting endpoints and model identifiers change between versions. Check them in the provider's current docs (`live-source-verification`) instead of writing them from memory. A wrong error-class name makes a fallback that never triggers.
5. **Verify**: run the evals, run the injection tests, simulate failures (mock the client to return timeouts, 429s, 5xx), and feed oversized inputs.

## Working rules

- **App type first.** Use the app profile from the audit if one exists, otherwise classify the project with `resources/app-profiles.md` in the `ship-readiness-audit` skill. Mark items N/A for the profile instead of reporting them MISSING, and use the profile's changed meaning where the table says so.
- **Change discipline.** Use `reviewable-diffs` if installed (workspace `.agents/skills/` or global `~/.gemini/config/skills/`). Otherwise: one approved item per branch or commit; change only the lines that item needs; no formatting, renaming or refactoring churn; no features nobody asked for; add a test that fails without the fix and passes with it; list each file touched and why.
- **Current facts.** Use `live-source-verification` if installed. Either way, apply the citation rule: any vendor-specific or time-sensitive fact (pricing, plan limits, config keys, CLI flags, SDK and error names, DNS values, legal or store rules) carries an inline [Title](URL) retrieved this session. No URLs from memory. If a source is unreachable, write "Source unavailable" and tag the claim [Unknown].
- **Confidence tags.** Tag each finding's evidence with exactly one of [Cited], [Certain], [Likely], [Guessing], [Unknown] ([Cited]: checked against a live source this session, with the link; [Certain]: seen directly this session; [Likely]: strong inference; [Guessing]: hypothesis only; [Unknown]: name the gap and what would close it). PASS requires [Certain] or [Cited].
- **Irreversible steps need their own yes**, even for approved items: rewriting git history, rotating or revoking live keys, DNS changes, production deploys, restoring or deleting data, billing changes.

## Item playbooks

Detailed guidance for each item is in `resources/items.md`. Read only the sections for the items in scope, plus the **Current requirements** section at the end, which holds externally verified standards and dates:

- A1 — Evals
- A2 — Prompt-injection testing
- A3 — Fallback when the model API goes down
- A4 — Token limits and context overflow
