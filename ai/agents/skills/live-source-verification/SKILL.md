---
name: live-source-verification
description: Verifies claims against live sources instead of memory. Use when a claim depends on information that must be checked live rather than recalled — current state (prices, versions, releases, who holds a role, what exists now), anything after the training cutoff, specific numbers, dates, named studies, direct quotes, any URL, or a factual claim the user will act on where being wrong is costly. Use whenever an answer would otherwise rest on remembered figures or links, even if the user doesn't ask for sources. Enforces atomic decomposition before verification, primary sources over aggregators, inline citation on the specific claim, attributed presentation of conflicting sources, a stopping rule, and reporting what a failed search found. Do NOT use for judgment, design opinion, reasoning about the user's own situation, or stable technical knowledge.
---

# Live-Source Verification

## 0. Tool check
This skill needs a web search or fetch tool (or browser use). If none is available in this session, say so up front, mark the affected claims [Unknown], and don't fill the gap from memory.

## 1. Decompose before verifying
Break the assertion into atomic checkable units. Verify each separately — a conclusion is only as sound as its weakest input. Name the input that couldn't be verified.

## 2. Primary sources first
Official docs, primary research, filings, changelogs, and vendor releases over aggregators, blogspam, and SEO content. Name the source type when it's weak.

## 3. Cite inline
Format: [Source Title](URL), placed on the claim it supports. Never batched at the end.

## 4. Conflicting sources
Present the disagreement with attribution. Say which is more authoritative and why. Don't average them into a false consensus.

## 5. Don't overextend a citation
The link must support the specific claim attached to it, not a broader version of it. If the source is adjacent rather than direct, say so.

## 6. Stopping rule
Verify until the load-bearing claims are covered, then stop. Don't keep searching for confirmation of something already established, and don't pad with sources that repeat the first one. If two searches on a specific point return nothing usable, say so and mark it [Unknown] rather than trying a third phrasing.

## 7. Report the search, not just the result
When verification comes up short, say what was looked for and what was found instead. "No official figure published; the number circulating comes from [aggregator]" beats both a confident number and silence. Never let an unverifiable claim silently drop out of the answer.

---
Tier assignment follows the user's always-on epistemic baseline if one is installed as a rule (`GEMINI.md`, `AGENTS.md`, or `.agents/rules/`). If none is installed, use these tiers on load-bearing claims only: [Cited] verified against a live source this session, with an inline [Title](URL); [Certain] stable and not time-sensitive; [Likely] strong inference; [Guessing] a hypothesis; [Unknown] can't find out, so name the gap and what would close it. A verified claim is [Cited]; an unreachable source drops to [Unknown].
