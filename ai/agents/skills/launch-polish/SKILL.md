---
name: launch-polish
description: Audits and then, with approval, finishes the launch-facing parts of a web app in any stack — SEO and AEO (robots.txt, sitemap, meta tags, canonical URLs, structured data), OG images and favicon, image optimization, real 404/500/empty/loading states, accessibility, analytics, custom domain and SSL, privacy policy and terms, cookie consent, account and data deletion, and a README. Use when someone is preparing a public launch and asks about search visibility, showing up in AI answers, link previews on social, favicons, slow images, error pages, blank screens, a11y or WCAG, tracking or analytics, connecting a domain or HTTPS, legal pages, GDPR-style consent banners, "delete my account", or documenting the project — even if they just say "make it feel finished". Part of the ship-readiness family (items P1–P11).
---

# Launch Polish and Compliance

These are the things users, search engines and regulators notice first: what a shared link looks like, what happens on a broken URL, whether a screen reader can use it, whether there's a privacy policy that tells the truth. Audit, report, implement only what's approved.

## Workflow

1. **Detect the stack**, rendering mode (server-rendered, static, client-only SPA), hosting, and every third-party script and SDK. The tracker inventory feeds P6, P8 and P9. If `ship-readiness-audit` is installed (workspace `.agents/skills/` or global `~/.gemini/config/skills/`), use its `scripts/triage_scan.py` (run it with `--help` first).
2. **Audit P1–P11** with PASS / PARTIAL / MISSING / UNKNOWN / N/A and evidence. When a URL is available, fetch real pages and headers. Rendered output beats reading source.
3. **Report and stop.** Use the ship-readiness table. Ask what to implement.
4. **Implement approved items** one per branch or commit (`reviewable-diffs` if installed). Check search-engine guidelines, schema.org types, accessibility criteria and privacy rules in current primary sources (`live-source-verification`). These change, and remembered rules go stale.
5. **Verify**: validate sitemaps and structured data, run an accessibility scanner, render pages at mobile width, curl redirects. Record the evidence.

Legal items (P8, P9, and parts of P10) have a hard limit. You can inventory data flows, draft documents, and implement mechanisms. You can't determine what's legally sufficient for the user's jurisdiction and business. Mark drafts "draft — not legal advice — have it reviewed", and say this once in the report rather than in every line.

## P1 — SEO and AEO

Technical SEO basics:
- `robots.txt` allows production pages, points to the sitemap, and doesn't accidentally `Disallow: /` (a common leftover from development). Staging and preview deployments are kept out of the index by a `noindex` header/meta or auth. Note that a robots disallow alone doesn't remove already-indexed pages.
- `sitemap.xml` lists canonical, indexable URLs only, is valid XML, and stays current (generated at build or runtime, not hand-maintained).
- Every indexable page has a unique `<title>` and meta description, one `<h1>`, and a `rel=canonical` to its preferred URL. There's one canonical host (http→https, and www vs apex; coordinate with P7).
- Important content and links are in the server-rendered HTML. Fetch the page without JS and look. Client-only rendering hurts crawlers that don't run JS and many AI crawlers.
- Structured data as JSON-LD, using schema.org types that match the content (Organization, WebSite, Product, Article, FAQPage, BreadcrumbList…). Validate it, and only mark up what's visible on the page.

AEO (answer-engine optimization, meaning being quotable by AI assistants and answer features) mostly overlaps with good SEO: crawlable server-rendered content, clear headings that state questions and answer them directly, concise factual summaries near the top, structured data, and pages that load fast. `llms.txt` is an emerging, unofficial convention. Adding it is cheap, but don't claim it affects ranking or citation. Ask whether the user wants AI crawlers allowed or blocked in `robots.txt`. It's a business choice.

## P2 — OG images and favicon

- Open Graph tags on shareable pages: `og:title`, `og:description`, `og:image` (absolute URL, commonly about 1200×630), `og:url`, `og:type`, plus Twitter/X card tags. Per-page images for key pages if the framework supports generating them.
- Favicon set: `favicon.ico` or SVG icon, `apple-touch-icon`, and a web manifest with icons if the app is installable. Check that the files are actually served, not just referenced.
- Verify by fetching the page HTML and confirming tags resolve to absolute, reachable URLs. Preview-card debuggers on social platforms are a Human check.

## P3 — Optimize images

- Serve modern formats (WebP/AVIF) with fallbacks where needed, via the framework's image component, an image CDN, or a build step.
- Responsive sizes (`srcset`/`sizes`) so phones don't download desktop images.
- Explicit `width`/`height` or aspect-ratio to prevent layout shift.
- `loading="lazy"` below the fold. Don't lazy-load the main above-the-fold (LCP) image, and consider prioritizing it.
- Flag any image over a few hundred KB served to clients, with its path and size.

## P4 — Real 404, 500, empty and loading states

- **404**: a custom page that returns an actual 404 status (not a 200 "soft 404"), keeps site navigation, and offers a way forward (home, search).
- **500 / error boundary**: a friendly page that doesn't leak stack traces, gives a way to retry or go home, and reports to error tracking (`ops-reliability` O2). Check both server errors and client-side render errors. Frameworks usually have separate hooks for each.
- **Empty states**: every list, table, dashboard and search result has a designed zero-data state that explains what goes here and how to add the first item. A blank area looks broken.
- **Loading states**: skeletons or spinners for data views, disabled buttons during submission, no layout jump when data arrives. Test with a throttled network (`test-and-ci` T7).
- **Error states per component**: a failed fetch shows a message and retry, not an infinite spinner.

Audit by listing every data-fetching component and marking which of loading, empty and error it handles.

## P5 — Make it accessible

Target WCAG 2.2 Level AA unless the user specifies otherwise. Check the current criteria from the W3C source when precision matters.

- **Automated scan** (axe-core or similar) of key pages and states, in e2e tests if possible. Automated tools catch only part of real issues, so a clean scan isn't a PASS by itself.
- **Keyboard pass**: every interactive element is reachable and operable by keyboard in a sensible order. Focus is visible. Modals trap and restore focus. No keyboard traps. There's a skip link on content-heavy pages.
- **Semantics**: real buttons and links (not clickable divs), labeled form fields with programmatic error messages, landmarks, heading order, `lang` attribute, meaningful alt text (empty alt for decorative images).
- **Visual**: text contrast, no information by color alone, usable at 200% zoom and 320px width, respects reduced-motion preferences.
- **Screen reader spot-check** of one core flow is a Human task. Write the steps.

Report issues grouped by page, with severity and the exact element.

## P6 — Create analytics

Analytics should answer the questions the user actually has. Ask for 3–5 product questions first ("Do people who sign up finish onboarding?"), then define events to answer them: a clear name, when it fires, properties, and **no personal data in properties** unless deliberately designed. Write a short tracking plan in the repo (`docs/analytics.md`).

Implement with the tool the user picks (privacy-friendly, cookieless options exist and can reduce consent needs; that's a factor in P9). Fire key events server-side where accuracy matters (signup, purchase), since ad blockers drop client events. Gate on consent where required (P9). Verify events arrive in the tool's debug view.

## P7 — Custom domain and SSL

Mostly Human (DNS registrar and host dashboard). Audit with:
- `curl -sI http://<domain>`: expect a redirect to https.
- `curl -sI https://<domain>` and `https://www.<domain>`: one redirects to the other (the canonical host). Both have valid certificates.
- Certificate auto-renewal is managed by the host or a client (not a manually uploaded certificate that will expire). Note the expiry date.
- Consider HSTS once HTTPS is confirmed everywhere. It's hard to undo, so stage it, starting with a short max-age.
- Get the exact DNS records from the host's current docs for the user's setup (apex vs subdomain). Don't guess IPs or CNAME targets.

## P8 — Privacy policy and terms

Start from facts, not a template. Build a **data inventory** from the code: what personal data is collected (forms, auth, analytics, logs, uploads, LLM prompts), where it's stored, which processors receive it (host, DB provider, email, analytics, error tracking, payment, LLM provider), retention, and whether anything is sold or used for ads. The inventory is the valuable part, because an inaccurate policy is worse than a generic one.

Then draft the privacy policy and terms from the inventory, clearly marked as drafts for legal review. Link both in the footer and at signup, and anywhere data is collected. Note that policies must be updated when processors change, and suggest keeping the inventory in the repo next to the code that creates it.

## P9 — Cookie consent

First inventory every cookie, local-storage key and third-party script, and classify each as strictly necessary or not (analytics, ads, session replay, embeds that set cookies). If everything is strictly necessary, or analytics is cookieless and doesn't need consent in the user's markets, a banner may not be needed. That's a legal judgment, so present it as a question for the user, not a conclusion.

Where consent is required (the EU/UK-style regimes are the common case; requirements differ elsewhere), the mechanism needs to:
- **Block non-essential scripts and cookies until consent.** A banner that loads trackers anyway does nothing. Verify in a fresh browser profile that no non-essential requests fire before consent.
- Make accepting and rejecting comparably easy, use granular categories, store the choice, and let users change it later (footer link).
- Pass consent state to analytics/ads tools that support consent modes.

## P10 — Let users delete their account and data

An in-product "Delete account" flow (re-authentication or confirmation required) that triggers a deletion job covering:
- Primary database rows (delete or anonymize, with foreign keys and orphaned records handled deliberately)
- Uploaded files in object storage
- The auth provider's user record
- Third-party processors: payment customer, email-list contact, analytics person profile, support tool, LLM provider stored data if any. Use each one's deletion API where available, otherwise document a manual step.
- Sessions revoked. Confirmation email sent.

Document what remains and for how long: backups that age out on their retention schedule, and records kept for legal or tax reasons. Put that in the privacy policy (P8). Some app stores require in-app account deletion for apps that allow account creation. If there's a mobile app, check the current store rules.

Tests: delete a test user and assert their rows, files and auth record are gone or anonymized, and that their old session no longer works.

## P11 — Write a README

Write for "future you" six months from now, under stress:
- What it is, in one paragraph, plus the architecture in a few lines (frontend, backend, DB, key services).
- Local setup: prerequisites with versions, install, env vars (names and purpose; values never, and point to `.env.example`), seed data, run.
- Testing: how to run unit, e2e and evals.
- Deploy: how a change reaches production, and where staging is.
- Operations: links to `ROLLBACK.md`, `RESTORE.md`, where logs, errors, uptime and billing dashboards live (names, not credentials), and who owns what.
- Known gotchas.

Extend an existing README rather than replacing it. Every command in it should be one you ran or verified in the repo. Mark anything unverified.
