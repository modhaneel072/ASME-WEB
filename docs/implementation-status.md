# Implementation status

_Last updated: 2026-09-08 (Stage 0 → Stage 2 in one run)._ Everything marked **Done** is persisted, authorised on the server, and covered by a test named below. Nothing in navigation is decorative.

## Stage summary

| Stage | Scope | Status |
| --- | --- | --- |
| 0 – Baseline | ADR-0001, design references, permissions matrix, screen inventory | Done |
| 1 – Foundation | tenancy, roles/permissions/scopes, audit, typed API errors, tokens + UI primitives, app shell, Setup Center, teams/users, locations, categories | Done |
| 2 – Projects + Work Orders | projects (list/detail/milestones/members/files/activity), work orders (list/filters/saved views/To Do–Done, right-side create pane, detail with comments/files/history/assignees/time/costs/sub-work orders), assets, operations dashboard, seeded Crater Cruncher Rover data | Done |
| 3+ | requests, parts, procedures, maintenance plans, automations, more reporting, email notifications | Not started (listed as "Not in this release" in Setup Center; no navigation) |

## Gate results (run on this commit)

| Gate | Command | Result |
| --- | --- | --- |
| Backend tests | `ASME_ENV=testing python -m pytest --ignore=tests/landing` | **222 passed** (155 legacy + 67 ops incl. migrations, seeds, dashboard) |
| Frontend typecheck | `npm run typecheck` | clean |
| Frontend lint | `npm run lint` | 0 errors, 6 warnings (`react-refresh/only-export-components` on files that export helpers next to components) |
| Frontend unit | `npm run test` | **22 passed** (6 files) |
| Frontend build | `npm run build` | ok – `dist/` 4 chunks, 1.05 kB html, 38.7 kB css, JS 93 + 165 + 287 + 400 kB (gzip 26 + 54 + 76 + 109 kB) |
| End-to-end | `npm run test:e2e` (all four viewports) | **23 passed** – 17 desktop-1440 scenarios + `@responsive` layout and axe checks at 1366×768, 768×1024, 390×844 |
| `tests/landing` | excluded | pre-existing landing-page suite needs a browser toolchain not part of this work |

Playwright attaches screenshots (`work-orders-list-*`, `work-order-detail-*`, `work-order-create-*`) and axe JSON per viewport to `apps/ops-web/playwright-report/`.

## What was built (by area)

### Backend (`asme/ops`, migration `0003_ops_foundation`)
- Organisation tenancy, memberships, 12 system roles with scoped permissions (`docs/permissions-matrix.md`), object-level checks, cross-tenant ids → 404.
- Work orders: per-org numbering, DRAFT→OPEN→IN_PROGRESS→ON_HOLD/DONE/CANCELED lifecycle with history and audit, assignees (people + teams), watchers, categories, assets, sub-work orders with auto-complete policy, recurrence with idempotent generation, completion with note/time/costs/asset status/follow-up, duplicate, comments (edit/delete own), attachments with signed expiring links, time and cost entries, notifications.
- Projects with milestones, members (project role + team), health and activity; teams with leads; locations (tree, default); categories; assets (tree, status history, types).
- Setup Center evaluation from persisted rows; operations dashboard; saved filters; audit log; OpenAPI 3.1 document.
- Seeds: idempotent defaults at boot; `manage.py seed-demo` (development only, documented password); `manage.py serve-e2e`.

### Frontend (`apps/ops-web`)
- Tokens and primitives per spec §6–7 (sidebar 264px, 40px controls, 36px chips, 52px rows, right-side 560px create pane, dialogs, popovers, toasts, empty/skeleton/error states).
- Screens listed in `docs/reference/screen-inventory.md`; all list views are URL-backed (`filter[...]`, `sort`, `q`) and shareable.
- Every mutation goes through the typed API client with the CSRF header; server field errors land on the matching form field; 401 returns to login with `next`.

## Defects fixed while verifying
- Dashboard weekly series dropped the last days of a 30-day range (4×7 buckets) – last bucket now extends to the range end (`tests/ops/test_dashboard.py`).
- Demo seed conflicted with the ops project mirrored from the public "rover" project – it now adopts that row.
- Escape inside a side sheet closed the sheet while a picker was open – the open popover now owns Escape.
- Creating a work order left the pane open because two URL updates raced – single update.
- Mobile drawer did not close on Escape or when tapping the current route – fixed.
- Invited (pending) members were hidden from the Users page – shown by default; suspended/disabled behind the toggle.
- Recharts pie sectors carried `role="img"` without names (axe) – replaced with the accessible bar list; inline links underlined (WCAG 1.4.1).
- `tests/test_scheduling.py::test_double_booking_conflicts` looked one day ahead and failed after the last work-hour of the day – now looks two days ahead.

## Known gaps and unknowns (truthful)
- **Reference recording not available**: `docs/reference/video-observations.md` and `design-measurements.md` are INFERRED/PROPOSED, not measured from a recording.
- Email delivery of notifications, virus scanning of uploads (`scan_status = skipped`), organisation-timezone-aware dashboard buckets, calendar/workload layouts for work orders, custom fields UI and meters for assets, notification/integration settings – not implemented.
- Attachments are stored on local disk (instance folder); object storage is a later adapter.
- The legacy `members` table and legacy portal stay for one more release (see `docs/migration-plan.md`).
- `react-refresh` lint warnings are cosmetic; splitting helper exports into separate files is a follow-up.
- Prettier is applied to `src/`, `e2e/` and config files; Python code follows the existing repo style (no formatter enforced).

## How to resume

```bash
# backend
ASME_ENV=testing .venv/Scripts/python.exe -m pytest -q --ignore=tests/landing
python manage.py seed-demo && python manage.py serve       # http://127.0.0.1:5000/app  admin@uiowa.edu / ChangeMe123!

# frontend
cd apps/ops-web && npm ci && npm run build && npm run test:e2e
```

Next task: Stage 3 (requests portal, parts inventory bridge to the existing stock ledger, procedures library), starting with the request model and the `/app/requests` screen, following the same slice discipline (model → service → API → screen → tests → docs).
