# Changelog

## Unreleased – ASME Ops stage 1 + 2

### Added
- **ASME Ops** operations platform served at `/app` (React 18 + TypeScript, `apps/ops-web`) on top of the existing Flask app, behind `ASME_OPS_ENABLED`.
- Organisation-scoped domain (`asme/ops`): organisations, memberships, 12 system roles with scoped permissions, teams, locations, categories, assets with status history, projects with milestones and members, work orders with lifecycle/history/assignees/watchers/sub-work orders/recurrence/time/costs, comments, attachments with signed download links, notifications, saved filters, audit events.
- `/api/v1` ops API with typed error envelope, CSRF header requirement, `filter[...]`/`sort`/`page[cursor]` list conventions and an OpenAPI 3.1 document at `/api/v1/ops/openapi.json`.
- Screens: login, Setup Center, Reporting (operations dashboard), Work Orders (list/table, filters, saved views, right-side create pane, detail with comments/files/time/costs/history), Projects (cards, detail tabs), Teams & Users, Locations, Categories, Assets, Settings (chapter, roles matrix, audit log).
- Alembic migration `0003_ops_foundation` (additive only) and idempotent seeds; `python manage.py seed-demo` creates the Crater Cruncher Rover demo chapter (development only); `python manage.py serve-e2e` boots a throwaway seeded server for Playwright.
- Tests: ops pytest suite (authz, work orders, projects/directory, API, dashboard, seeds, migrations), vitest unit tests, Playwright end-to-end scenarios with responsive screenshots and axe checks.
- Docs: ADR-0001, architecture overview, API reference, permissions matrix, reporting metrics, migration plan, deployment, test plan, implementation status, design references.

### Changed
- `/app` now serves ASME Ops; the legacy portal moved to `/legacy/app` (unchanged otherwise, gated by `ASME_ENABLE_LEGACY_OPS`). `/auth/login` redirects to `/app/login`.
- Changing a member's ops role keeps the legacy `users.role` in sync so the legacy portal continues to authorise correctly.
- Request datetimes are normalised to naive UTC at the contract boundary.

### Fixed
- Operations dashboard weekly series dropped the last days of a range (30 days → 4×7 buckets); the final bucket now extends to the range end.
- Demo seed adopts the ops project that the default seed mirrors from the public "rover" project instead of failing with a conflict.
- OpenAPI paths are emitted relative to the `/api/v1` server URL.

### Not in this release (no navigation, listed in Setup Center as "Not in this release")
Requests, messages, events, parts/inventory, purchase requests, vendors, sponsors, procedures library, maintenance plans, meters, automations, additional reporting, email notifications, integrations settings.
