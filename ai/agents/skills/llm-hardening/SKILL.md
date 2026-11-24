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
3. **Report and stop.** Use the ship-readiness table. Ask what to implement.
4. **Implement approved items** one per branch or commit (`reviewable-diffs` if installed). SDK method names, error types, stop-reason values, token-counting endpoints and model identifiers change between versions. Check them in the provider's current docs (`live-source-verification`) instead of writing them from memory. A wrong error-class name makes a fallback that never triggers.
5. **Verify**: run the evals, run the injection tests, simulate failures (mock the client to return timeouts, 429s, 5xx), and feed oversized inputs.

## A1 — Evals

Evals are the unit tests of prompts. Without them, every prompt tweak or model upgrade is a guess.

Build a starter set:
- **Cases**: 20–50 representative inputs per feature to start, drawn from real usage if available (anonymized), plus hand-written edge cases: empty input, very long input, other languages, ambiguous requests, adversarial-but-plausible inputs, and inputs that should be refused or deflected. Store them in the repo (e.g. `evals/<feature>/cases.jsonl`), versioned.
- **Graders**, cheapest first: exact or structural checks (valid JSON, required fields, contains or doesn't contain), programmatic checks (numbers within tolerance, links valid), then model-graded rubrics for subjective quality. Model graders need their own spot-check against human judgment, or they quietly drift.
- **A runner** that reports a pass rate per feature and per grader, and diffs against the last baseline.
- **When to run**: before any change to prompts, model, temperature, retrieval or tools, and in CI for changes touching those files (cost-capped, perhaps a smaller subset). Set the pass threshold with the user. It's a product decision.

Ask the user what "good" means for each feature before writing rubrics. Rubrics invented without them measure the wrong thing.

## A2 — Prompt-injection testing

Any text the model reads can contain instructions: user messages (direct injection) and, more dangerously, content the app feeds in from elsewhere (indirect injection), such as retrieved documents, web pages, emails, uploaded files, tool or API results, and earlier conversation. The question isn't "can the model be tricked". Assume it sometimes can. The question is "what's the worst that happens when it is?"

Audit the design first, since that's where most of the protection lives:
- **Least privilege for tools.** What can the model do without a human confirming? Sending email, spending money, deleting data, calling internal APIs with broad credentials, or reading other users' data are high-risk if reachable from untrusted content. Prefer scoped credentials per user, and confirmation steps for consequential actions.
- **Data the model can see.** Secrets, other users' data, or internal-only info in the system prompt or context can leak. Assume the system prompt can be extracted. Don't put secrets in it.
- **Output handling.** Never execute model output, and never pass it unvalidated into SQL, shell, file paths or URLs. Don't render it as raw HTML. Treat it like user input, escape it, and sanitize markdown links and images (image URLs can exfiltrate data via query strings).
- **Separation.** Wrap untrusted content in clear delimiters and tell the model it's data, not instructions. This helps but isn't a guarantee, so don't rely on it alone.

Then test: write a suite (it can live alongside the evals) of injection attempts placed in each untrusted channel the app actually has. Examples: instructions to ignore prior rules, to reveal the system prompt, to call a tool with attacker-chosen arguments, to emit a link or image to an external URL with data in it, and to change the output format to break the parser. Grade on outcomes ("no tool call made", "no external URL in output", "JSON still valid"), not on the model's wording. These are tests of the user's own system. Keep payloads generic and focused on the app's real channels and tools.

Report residual risk honestly. Some injection will succeed against current models, so the mitigation that matters is limiting what a successful injection can do.

## A3 — Fallback when the model API goes down

Wrap model calls in one client module so resilience logic lives in one place. Check or add:
- **Timeouts** on every call, set per use case. Streaming calls need a time-to-first-token timeout and an idle timeout.
- **Retries with exponential backoff and jitter**, only on retryable conditions (rate limits, overloaded, 5xx, network errors), respecting any retry-after hint, with a small max attempt count. Never retry on invalid-request errors. Make sure retries aren't stacked (SDK retries plus your own retries plus a client retry multiply).
- **Circuit breaker**: after repeated failures, stop calling for a short period and serve the fallback immediately, instead of making every user wait for timeouts.
- **Fallback path**, in order of preference for that feature: a secondary model or provider (needs its own evals, since prompts don't transfer perfectly), a cached or precomputed answer, a reduced non-AI version of the feature, or a clear user-facing message with retry. Pick per feature with the user.
- **Idempotency** for actions triggered by model output, so a retry doesn't do the thing twice.
- **Observability**: log provider, model, latency, token usage, and error type per call (`ops-reliability` O1/O2), and alert on failure-rate spikes (O3).

Verify by mocking the client to time out, return 429, return 5xx and drop the stream mid-way. Assert the user sees the degraded state, not a crash or infinite spinner.

## A4 — Token limits and context overflow

Two separate limits matter: the **context window** (input plus output) and the **output cap** (`max_tokens` or the equivalent).

Input side:
- Measure before sending. Use the provider's token-counting method or a tokenizer that matches the model, or at least a conservative character-based estimate with a safety margin. Check which the provider currently offers.
- Define a budget per call: system prompt, retrieved context, conversation history, user input, and reserved output.
- Have a strategy when over budget, chosen per feature: drop oldest turns (keeping the system prompt and recent turns), summarize older history, retrieve only top-k relevant chunks, chunk large documents and map-reduce, or reject with a clear message ("this file is too large, try a section"). Never silently cut the user's current message.
- Cap user-supplied input size at the API boundary. This also limits cost and abuse (`security-hardening` S4, `ops-reliability` O11).

Output side:
- Detect truncation. Check the response's stop/finish reason for the max-tokens case, using the provider's current field name and value. A truncated JSON response that the app then fails to parse is a common production bug.
- Handle it: raise the cap if appropriate, ask for continuation, or return a graceful error.

Tests: feed inputs just under and over the budget, and a response mocked as truncated. Assert the defined behavior for each.
