# Test plan

Three layers, all runnable locally and in CI. Exact commands and the latest results are in [implementation-status.md](implementation-status.md).

## 1. Backend (pytest)

```bash
ASME_ENV=testing python -m pytest -q --ignore=tests/landing
```

| Area | Files | What is pinned |
| --- | --- | --- |
| Legacy platform | `tests/test_*.py` (155 tests from the backend rewrite) | public site, portal, inventory ledger, prints, calendar, Launchpad, API v1, migrations |
| Authorisation | `tests/ops/test_authz.py` | permission catalogue, role grants, scope resolution, object-level checks, tenant isolation (`get_or_404`) |
| Work orders | `tests/ops/test_work_orders.py` | per-org numbering, lifecycle transitions and invalid ones, drafts, cancel reasons, completion aggregates (time/costs/asset status/follow-up), recurrence generated exactly once, sub-work orders + auto-complete policy, duplicate, critical priority gating, due-before-start, list tabs/filters/sorts/search, cursor paging, tab counts, time entries on behalf, comment notifications |
| Projects & directory | `tests/ops/test_projects_directory.py` | project CRUD/archive/members/milestones/health/activity, teams with leads, locations (default, nesting, archive), categories (unique names) |
| API surface | `tests/ops/test_api_ops.py` | login/session, CSRF, error envelope, validation details, 401/403/404 shapes, list conventions, file upload/download token expiry, saved filters, OpenAPI document |
| Dashboard | `tests/ops/test_dashboard.py` | every formula in `docs/reporting-metrics.md` |
| Seeds | `tests/ops/test_seeds.py` | defaults idempotent; demo seed idempotent, coherent, one DRAFT, mirrored public project adopted |
| Migrations | `tests/ops/test_migrations.py` | fresh database and a populated 0002 database both upgrade to head with backfilled memberships; downgrade path |

## 2. Frontend unit (vitest + Testing Library)

```bash
cd apps/ops-web && npm run test
```

`api/client.test.ts` (envelope, CSRF header, typed errors, 401 listener, offline), `contracts/schemas.test.ts` (form → payload normalisation, date validation), `lib/format.test.ts`, `lib/listParams.test.tsx` (URL ↔ list state), `ui/Badge.test.tsx`, `ui/FilterChip.test.tsx` (open, multi-select, clear, Escape).

## 3. End-to-end (Playwright against the built app)

```bash
cd apps/ops-web && npm run build && npm run test:e2e            # all projects
npx playwright test --project=desktop-1440                        # one viewport
```

`playwright.config.ts` starts `python manage.py serve-e2e` (fresh SQLite in the temp folder, demo seed, port 5055) and serves `apps/ops-web/dist`. Scenarios:

| Spec | Scenario |
| --- | --- |
| `auth.spec.ts` | redirect to login with `next`, sign in by email and by username, wrong password message, legacy/public/alias routes still respond |
| `projects.spec.ts` | admin creates a project and assigns a lead, adds a milestone, sees members; project work-orders tab is scoped and presets the project in the create pane |
| `work-orders.spec.ts` | team lead creates a work order in the right-side pane (geometry asserted) assigning a member; the member filters "Assigned to me", starts it, comments, uploads a file, logs time, completes with a note, sees it under Done; sub-work orders, saved views, table layout |
| `authorization.spec.ts` | member sees no create-project button or admin nav, API returns 403 `forbidden` for `POST /projects`, 403 `csrf` without the header, gated page shows an access message, audit log 403, anonymous 401, OpenAPI public |
| `setup-and-manage.spec.ts` | Setup Center reflects persisted rows and shows "Not in this release" with no CTA for unbuilt modules; sidebar contains only built modules; chapter profile save completes a task; guide read; teams/users/locations/categories CRUD including a server-side duplicate error; asset status change; dashboard numbers |
| `responsive.spec.ts` (`@responsive`, runs at 1440×900, 1366×768, 768×1024, 390×844) | drawer vs sidebar, list/detail vs stacked with back button, create pane, no horizontal overflow, screenshots attached to the report; axe (WCAG 2A/AA, serious+critical) on login, setup center, work orders, reporting |

Artefacts: `apps/ops-web/playwright-report/` (HTML, git-ignored) with screenshots and axe JSON attached per viewport.

## Gates before merging

`npm run typecheck`, `npm run lint`, `npm run test`, `npm run build`, `pytest`, `npm run test:e2e`. All six are recorded in implementation-status.md with their real output.
