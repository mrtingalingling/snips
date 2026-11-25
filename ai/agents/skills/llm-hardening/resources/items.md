# LLM Hardening: item playbooks

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

## Current requirements (verified November 2026)

These change. Before relying on one in a report or runbook, re-check the linked source and cite it.

- **A1–A4 Taxonomy.** Map coverage to the OWASP Top 10 for LLM Applications 2026, published August 2026 ([OWASP GenAI](https://genai.owasp.org/resource/owasp-genai-llm-top-10-2026/)):
  - LLM01 Prompt Injection
  - LLM02 Sensitive Information Disclosure
  - LLM03 Excessive Agency
  - LLM04 Supply Chain
  - LLM05 Data and Model Poisoning
  - LLM06 Unbounded Consumption
  - LLM07 Misinformation
  - LLM08 Hidden Context Exposure
  - LLM09 Vector and Embedding Weaknesses
  - LLM10 Improper Output Handling
- **A2 Excessive agency.** It now ranks #3. When the model has tools, memory or the ability to act, test that every tool call is least-privilege, and that side-effecting actions need confirmation outside the model. For agentic features, also use OWASP's separate Top 10 for Agentic Applications, released December 2025 ([Bugcrowd overview](https://www.bugcrowd.com/blog/owasp-top-10-security-threats-facing-ai-systems/)).
- **A2 Hidden context exposure (LLM08).** This was "System Prompt Leakage" in 2025. Assume system prompts, tool schemas and retrieved context can be extracted. Keep secrets and authorization logic out of prompts and in code.
- **A3/A4 Unbounded consumption (LLM06).** Per-user token and cost budgets (O11) and input-size limits (A4) are the controls for this category.
