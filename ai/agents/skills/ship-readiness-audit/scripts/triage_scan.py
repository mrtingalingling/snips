#!/usr/bin/env python3
"""
triage_scan.py — emit production-readiness *signals* for a repo as JSON.

Signals are hints about where to look, never verdicts. Standard library only.
Scans the working tree, not git history; secrets that were committed and later
removed need a history scanner.

Usage: python3 triage_scan.py [repo_root] > triage.json
"""
import argparse
import json
import os
import re
import subprocess
import sys
from pathlib import Path

SKIP_DIRS = {
    ".git", "node_modules", ".next", "dist", "build", "out", ".venv", "venv",
    "__pycache__", ".turbo", ".cache", "coverage", "vendor", "target", ".svelte-kit",
    ".nuxt", ".output", ".vercel", ".terraform",
}
MAX_FILE_BYTES = 1_000_000

# Dependency name -> (category, label). Matching is on dependency names in manifests.
DEP_SIGNALS = {
    # frameworks
    "next": ("framework", "Next.js"), "react": ("framework", "React"),
    "vue": ("framework", "Vue"), "nuxt": ("framework", "Nuxt"),
    "svelte": ("framework", "Svelte"), "@sveltejs/kit": ("framework", "SvelteKit"),
    "astro": ("framework", "Astro"), "@remix-run/node": ("framework", "Remix"),
    "express": ("framework", "Express"), "fastify": ("framework", "Fastify"),
    "@nestjs/core": ("framework", "NestJS"), "hono": ("framework", "Hono"),
    "django": ("framework", "Django"), "flask": ("framework", "Flask"),
    "fastapi": ("framework", "FastAPI"), "rails": ("framework", "Rails"),
    "laravel/framework": ("framework", "Laravel"),
    # tests
    "jest": ("tests", "Jest"), "vitest": ("tests", "Vitest"), "mocha": ("tests", "Mocha"),
    "@playwright/test": ("tests", "Playwright"), "cypress": ("tests", "Cypress"),
    "pytest": ("tests", "pytest"), "rspec": ("tests", "RSpec"),
    # observability
    "@sentry/node": ("error_tracking", "Sentry"), "@sentry/nextjs": ("error_tracking", "Sentry"),
    "@sentry/react": ("error_tracking", "Sentry"), "@sentry/browser": ("error_tracking", "Sentry"),
    "sentry-sdk": ("error_tracking", "Sentry"), "@bugsnag/js": ("error_tracking", "Bugsnag"),
    "@honeybadger-io/js": ("error_tracking", "Honeybadger"), "rollbar": ("error_tracking", "Rollbar"),
    "pino": ("logging", "pino"), "winston": ("logging", "winston"),
    "structlog": ("logging", "structlog"), "loguru": ("logging", "loguru"),
    "@opentelemetry/api": ("perf", "OpenTelemetry"), "opentelemetry-api": ("perf", "OpenTelemetry"),
    "dd-trace": ("perf", "Datadog APM"), "newrelic": ("perf", "New Relic"),
    "web-vitals": ("perf", "web-vitals"),
    # analytics
    "posthog-js": ("analytics", "PostHog"), "posthog-node": ("analytics", "PostHog"),
    "@vercel/analytics": ("analytics", "Vercel Analytics"), "mixpanel-browser": ("analytics", "Mixpanel"),
    "@amplitude/analytics-browser": ("analytics", "Amplitude"), "plausible-tracker": ("analytics", "Plausible"),
    # LLM
    "openai": ("llm", "OpenAI SDK"), "@anthropic-ai/sdk": ("llm", "Anthropic SDK"),
    "anthropic": ("llm", "Anthropic SDK"), "ai": ("llm", "Vercel AI SDK"),
    "langchain": ("llm", "LangChain"), "@langchain/core": ("llm", "LangChain"),
    "@google/generative-ai": ("llm", "Google GenAI SDK"), "google-genai": ("llm", "Google GenAI SDK"),
    "llamaindex": ("llm", "LlamaIndex"), "llama-index": ("llm", "LlamaIndex"),
    # auth / data / rate limit / email
    "next-auth": ("auth", "NextAuth/Auth.js"), "@clerk/nextjs": ("auth", "Clerk"),
    "@supabase/supabase-js": ("auth", "Supabase"), "firebase": ("auth", "Firebase"),
    "passport": ("auth", "Passport"), "lucia": ("auth", "Lucia"),
    "prisma": ("db", "Prisma"), "@prisma/client": ("db", "Prisma"), "drizzle-orm": ("db", "Drizzle"),
    "mongoose": ("db", "Mongoose"), "sqlalchemy": ("db", "SQLAlchemy"),
    "express-rate-limit": ("rate_limit", "express-rate-limit"),
    "@upstash/ratelimit": ("rate_limit", "Upstash Ratelimit"), "slowapi": ("rate_limit", "slowapi"),
    "django-ratelimit": ("rate_limit", "django-ratelimit"), "rack-attack": ("rate_limit", "rack-attack"),
    "resend": ("email", "Resend"), "@sendgrid/mail": ("email", "SendGrid"),
    "nodemailer": ("email", "Nodemailer"), "postmark": ("email", "Postmark"),
    "stripe": ("billing", "Stripe"),
}

INFRA_FILES = {
    "vercel.json": "Vercel", "netlify.toml": "Netlify", "fly.toml": "Fly.io",
    "render.yaml": "Render", "railway.json": "Railway", "wrangler.toml": "Cloudflare Workers",
    "Dockerfile": "Docker", "docker-compose.yml": "Docker Compose", "compose.yaml": "Docker Compose",
    "serverless.yml": "Serverless Framework", "app.yaml": "Google App Engine", "Procfile": "Heroku-style Procfile",
    "firebase.json": "Firebase Hosting", "amplify.yml": "AWS Amplify",
}
CI_PATHS = [".github/workflows", ".gitlab-ci.yml", ".circleci", "azure-pipelines.yml",
            "bitbucket-pipelines.yml", "Jenkinsfile", ".buildkite"]

PRESENCE = {
    "robots_txt": ["robots.txt", "public/robots.txt", "static/robots.txt", "app/robots.ts", "app/robots.js"],
    "sitemap": ["sitemap.xml", "public/sitemap.xml", "static/sitemap.xml", "app/sitemap.ts", "app/sitemap.js"],
    "favicon": ["favicon.ico", "public/favicon.ico", "static/favicon.ico", "app/favicon.ico", "app/icon.png"],
    "readme": ["README.md", "README", "readme.md", "README.rst"],
    "env_example": [".env.example", ".env.sample", ".env.template"],
    "rollback_doc": ["ROLLBACK.md", "docs/ROLLBACK.md", "docs/rollback.md"],
    "restore_doc": ["RESTORE.md", "docs/RESTORE.md", "docs/restore.md"],
    "privacy_page_hint": ["app/privacy", "pages/privacy.tsx", "pages/privacy.js", "src/pages/privacy.astro", "privacy.html"],
    "terms_page_hint": ["app/terms", "pages/terms.tsx", "pages/terms.js", "src/pages/terms.astro", "terms.html"],
    "not_found_page_hint": ["app/not-found.tsx", "app/not-found.js", "pages/404.tsx", "pages/404.js", "404.html", "src/pages/404.astro"],
    "error_page_hint": ["app/error.tsx", "app/global-error.tsx", "pages/500.tsx", "pages/_error.tsx", "500.html"],
}

SECRET_PATTERNS = {
    "aws_access_key_id": r"AKIA[0-9A-Z]{16}",
    "anthropic_key": r"sk-ant-[A-Za-z0-9_\-]{20,}",
    "openai_style_key": r"sk-(?!ant-)(?:proj-)?[A-Za-z0-9_\-]{32,}",
    "stripe_live_secret": r"sk_live_[A-Za-z0-9]{16,}",
    "github_token": r"gh[pousr]_[A-Za-z0-9]{30,}",
    "slack_token": r"xox[abprs]-[A-Za-z0-9\-]{10,}",
    "google_api_key": r"AIza[0-9A-Za-z_\-]{35}",
    "private_key_block": r"-----BEGIN (?:RSA |EC |OPENSSH |DSA )?PRIVATE KEY-----",
}
SECRET_RE = {k: re.compile(v) for k, v in SECRET_PATTERNS.items()}
TEXT_EXT = {".js", ".jsx", ".ts", ".tsx", ".mjs", ".cjs", ".py", ".rb", ".go", ".rs", ".php",
            ".java", ".kt", ".json", ".yml", ".yaml", ".toml", ".env", ".ini", ".cfg", ".md",
            ".txt", ".sh", ".html", ".vue", ".svelte", ".astro", ".tf", ".properties", ""}


def walk(root: Path):
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS]
        for f in filenames:
            yield Path(dirpath) / f


def read_deps(root: Path):
    deps = set()
    for pj in root.rglob("package.json"):
        if any(p in SKIP_DIRS for p in pj.parts):
            continue
        try:
            data = json.loads(pj.read_text(encoding="utf-8"))
            for key in ("dependencies", "devDependencies", "peerDependencies"):
                deps.update((data.get(key) or {}).keys())
        except Exception:
            pass
    for name in ("requirements.txt", "requirements-dev.txt", "pyproject.toml", "Pipfile",
                 "Gemfile", "composer.json", "go.mod", "Cargo.toml"):
        for f in root.rglob(name):
            if any(p in SKIP_DIRS for p in f.parts):
                continue
            try:
                text = f.read_text(encoding="utf-8", errors="ignore").lower()
            except Exception:
                continue
            for dep in DEP_SIGNALS:
                if re.search(r'(^|[\s"\'/])' + re.escape(dep.lower()) + r'($|[\s"\'=<>~\[,])', text, re.M):
                    deps.add(dep)
    return deps


def languages(root: Path):
    markers = {"package.json": "JavaScript/TypeScript", "pyproject.toml": "Python",
               "requirements.txt": "Python", "Gemfile": "Ruby", "go.mod": "Go",
               "Cargo.toml": "Rust", "composer.json": "PHP", "pom.xml": "Java",
               "build.gradle": "JVM (Gradle)"}
    return sorted({lang for m, lang in markers.items() if (root / m).exists()})


def git_tracked_env(root: Path):
    try:
        out = subprocess.run(["git", "-C", str(root), "ls-files"], capture_output=True,
                             text=True, timeout=20)
        if out.returncode != 0:
            return None
        files = out.stdout.splitlines()
        return [f for f in files if re.search(r"(^|/)\.env(\.|$)", f)
                and not re.search(r"\.(example|sample|template)$", f)]
    except Exception:
        return None


def redact(s: str) -> str:
    return s[:6] + "…" + f"[{len(s)} chars]"


def secret_hits(root: Path, limit=50):
    hits = []
    for f in walk(root):
        if f.suffix.lower() not in TEXT_EXT and not f.name.startswith(".env"):
            continue
        if f.name.endswith((".example", ".sample", ".template")):
            continue
        try:
            if f.stat().st_size > MAX_FILE_BYTES:
                continue
            text = f.read_text(encoding="utf-8", errors="ignore")
        except Exception:
            continue
        for i, line in enumerate(text.splitlines(), 1):
            for kind, rx in SECRET_RE.items():
                m = rx.search(line)
                if m:
                    hits.append({"file": str(f.relative_to(root)), "line": i,
                                 "kind": kind, "match": redact(m.group(0))})
                    if len(hits) >= limit:
                        return hits, True
    return hits, False


def main():
    ap = argparse.ArgumentParser(
        description="Emit production-readiness signals for a repo as JSON on stdout. "
                    "Signals are hints, not verdicts. Scans the working tree only "
                    "(not git history). Secret matches are redacted.")
    ap.add_argument("repo_root", nargs="?", default=".", help="repository root to scan (default: .)")
    args = ap.parse_args()
    root = Path(args.repo_root).resolve()
    deps = read_deps(root)
    found = {}
    for dep in sorted(deps):
        if dep in DEP_SIGNALS:
            cat, label = DEP_SIGNALS[dep]
            found.setdefault(cat, set()).add(label)
    infra = sorted({label for fn, label in INFRA_FILES.items() if (root / fn).exists()})
    ci = [p for p in CI_PATHS if (root / p).exists()]
    presence = {k: [p for p in paths if (root / p).exists()] for k, paths in PRESENCE.items()}
    hits, truncated = secret_hits(root)
    report = {
        "root": str(root),
        "note": "Signals only. Confirm each with evidence before marking PASS. Working tree only; scan git history separately for secrets.",
        "languages": languages(root),
        "detected": {k: sorted(v) for k, v in sorted(found.items())},
        "infra_config": infra,
        "ci_config": ci,
        "presence": presence,
        "git_tracked_env_files": git_tracked_env(root),
        "secret_pattern_hits": hits,
        "secret_hits_truncated": truncated,
    }
    json.dump(report, sys.stdout, indent=2)
    sys.stdout.write("\n")


if __name__ == "__main__":
    main()
