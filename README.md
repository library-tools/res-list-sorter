# Reservation List Sorter

## Weekly Security Audit

This repo runs a weekly dependency security check in `.github/workflows/security-audit.yml`.

- Schedule: every Monday at 08:00 UTC
- Blocking check: fails on `high`/`critical` production vulnerabilities
- Informational check: also reports full audit counts (including moderate/dev) in the workflow summary

If the weekly audit fails:

1. Open the failed run in **Actions** and review the `Block on high+ prod vulnerabilities` step output.
2. Run `npm audit --omit=dev --audit-level=high` locally to reproduce.
3. Apply updates via Dependabot PRs or run `npm audit fix` and retest.
4. If an update is unavailable, document temporary risk acceptance in the related issue/PR and monitor for upstream fixes.
