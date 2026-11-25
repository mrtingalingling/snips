---
name: launch-polish
description: Audits and then, with approval, finishes the launch-facing parts of a web app in any stack — SEO and AEO (robots.txt, sitemap, meta tags, canonical URLs, structured data), OG images and favicon, image optimization, real 404/500/empty/loading states, accessibility, analytics, custom domain and SSL, privacy policy and terms, cookie consent, account and data deletion, and a README. Use when someone is preparing a public launch and asks about search visibility, showing up in AI answers, link previews on social, favicons, slow images, error pages, blank screens, a11y or WCAG, tracking or analytics, connecting a domain or HTTPS, legal pages, GDPR-style consent banners, "delete my account", or documenting the project — even if they just say "make it feel finished". Part of the ship-readiness family (items P1–P11).
---

# Launch Polish and Compliance

These are the things users, search engines and regulators notice first: what a shared link looks like, what happens on a broken URL, whether a screen reader can use it, whether there's a privacy policy that tells the truth. Audit, report, implement only what's approved.

## Workflow

1. **Detect the stack**, rendering mode (server-rendered, static, client-only SPA), hosting, and every third-party script and SDK. The tracker inventory feeds P6, P8 and P9. If the `ship-readiness-audit` skill is installed (workspace `.agents/skills/` or global `~/.gemini/config/skills/`), use the `triage_scan.py` script in its `scripts/` folder (run it with `--help` first).
2. **Audit P1–P11** with PASS / PARTIAL / MISSING / UNKNOWN / N/A and evidence. When a URL is available, fetch real pages and headers. Rendered output beats reading source.
3. **Report and stop.** Use the ship-readiness table. Ask what to implement. **Skip this step for items the user already approved in this conversation** (through `/ship-check` or the audit's Phase 4): go straight to step 4.
4. **Implement approved items** one per branch or commit (`reviewable-diffs` if installed). Check search-engine guidelines, schema.org types, accessibility criteria and privacy rules in current primary sources (`live-source-verification`). These change, and remembered rules go stale.
5. **Verify**: validate sitemaps and structured data, run an accessibility scanner, render pages at mobile width, curl redirects. Record the evidence.

Legal items (P8, P9, and parts of P10) have a hard limit. You can inventory data flows, draft documents, and implement mechanisms. You can't determine what's legally sufficient for the user's jurisdiction and business. Mark drafts "draft — not legal advice — have it reviewed", and say this once in the report rather than in every line.

## Working rules

- **App type first.** Use the app profile from the audit if one exists, otherwise classify the project with `resources/app-profiles.md` in the `ship-readiness-audit` skill. Mark items N/A for the profile instead of reporting them MISSING, and use the profile's changed meaning where the table says so.
- **Change discipline.** Use `reviewable-diffs` if installed (workspace `.agents/skills/` or global `~/.gemini/config/skills/`). Otherwise: one approved item per branch or commit; change only the lines that item needs; no formatting, renaming or refactoring churn; no features nobody asked for; add a test that fails without the fix and passes with it; list each file touched and why.
- **Current facts.** Use `live-source-verification` if installed. Either way, apply the citation rule: any vendor-specific or time-sensitive fact (pricing, plan limits, config keys, CLI flags, SDK and error names, DNS values, legal or store rules) carries an inline [Title](URL) retrieved this session. No URLs from memory. If a source is unreachable, write "Source unavailable" and tag the claim [Unknown].
- **Confidence tags.** Tag each finding's evidence with exactly one of [Cited], [Certain], [Likely], [Guessing], [Unknown] ([Cited]: checked against a live source this session, with the link; [Certain]: seen directly this session; [Likely]: strong inference; [Guessing]: hypothesis only; [Unknown]: name the gap and what would close it). PASS requires [Certain] or [Cited].
- **Irreversible steps need their own yes**, even for approved items: rewriting git history, rotating or revoking live keys, DNS changes, production deploys, restoring or deleting data, billing changes.

## Item playbooks

Detailed guidance for each item is in `resources/items.md`. Read only the sections for the items in scope, plus the **Current requirements** section at the end, which holds externally verified standards and dates:

- P1 — SEO and AEO
- P2 — OG images and favicon
- P3 — Optimize images
- P4 — Real 404, 500, empty and loading states
- P5 — Make it accessible
- P6 — Create analytics
- P7 — Custom domain and SSL
- P8 — Privacy policy and terms
- P9 — Cookie consent
- P10 — Let users delete their account and data
- P11 — Write a README
