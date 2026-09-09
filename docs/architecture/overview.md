# ASME Ops – architecture overview

_Status: implemented as of this commit. Companion documents: [ADR-0001](../decisions/ADR-0001-transitional-architecture.md), [api.md](../api.md), [permissions-matrix.md](../permissions-matrix.md), [migration-plan.md](../migration-plan.md)._

## Shape

```
browser ── /app/*  ──► React 18 SPA (apps/ops-web, built to dist/, served by Flask)
        ── /api/v1 ──► Flask blueprint `ops_api` (asme/ops/api) ── services ── SQLAlchemy models ── SQLite (dev) / PostgreSQL (prod)
        ── /, /projects/<slug>, /legacy/app, /kiosk, … ──► existing public site + legacy portal (unchanged)
```

One Python process serves everything. There is no second service to deploy; the frontend
build step is the only addition to the pipeline (see [deployment.md](../deployment.md)).

## Backend layout (`asme/ops`)

| Path | Responsibility |
| --- | --- |
| `models/` | Organisation-scoped tables. Every ops table has a UUID `id`, `organization_id` and audit columns (`OpsBase`). Work orders (`work.py`), projects and milestones (`projects.py`), assets (`assets.py`), everything else (`core.py`). |
| `authz.py` | Permission catalogue, the 12 system roles, scope resolution (`chapter > project > team > assigned > own`), `can()` / `require()` and the `permission_required` decorator. |
| `tenancy.py` | Resolves the organisation for a request, loads the `AuthzContext`, and provides `scoped()` / `get_or_404()` so cross-tenant ids look like 404s. |
| `services/` | All business rules. Views never touch the ORM directly. `work_orders.py` owns the lifecycle, numbering, recurrence and notifications; `projects.py`, `directory.py` (teams, locations, categories), `assets.py`, `organizations.py`, `setup_center.py`, `dashboard.py`, `comments.py`, `files.py`, `saved_filters.py`, `notifications.py`, `audit_log.py`. |
| `api/` | Thin Flask views: parse (`pydantic`), authorise, call a service, serialise. `query.py` implements the `filter[...]` / `sort` / `page[cursor]` conventions. `openapi.py` builds the OpenAPI 3.1 document from the `@contract` annotations. |
| `schemas/` | Pydantic request/response contracts. Request datetimes are normalised to naive UTC (`UtcDateTime`). |
| `storage.py` | Local attachment storage with signed, expiring download URLs (`itsdangerous`). |
| `audit.py` | `record_event()` writes before/after snapshots to `ops_audit_events`. |
| `seeds.py` | `seed_ops_defaults()` (roles, permissions, default org, default categories/asset types) and `seed_ops_demo()` (Crater Cruncher Rover demo data, development only). |
| `web.py` | Serves the built SPA under `/app` and redirects `/auth/login` to `/app/login`. |

### Request lifecycle

1. `ops_api.before_request` rejects state-changing requests without `X-Requested-With: ASME-Ops` (CSRF), then resolves the signed-session user and loads the `AuthzContext` (`g.ops_ctx`). Public endpoints: login, health, OpenAPI, signed file download.
2. The view validates the JSON body against a pydantic model; validation errors become `{ok:false, code:"validation", details:[{field,message}]}` with HTTP 400.
3. Authorisation happens in the view or service via `authz.require(...)` with the record's scope. Failures return 403 with `code:"forbidden"` and the missing permission.
4. Services mutate, write audit events and notifications, and commit.
5. Responses are wrapped as `{ok:true, payload:…}`. Lists return `{items, next_cursor}`.

### Data model highlights

- `organizations` → `memberships` (user × org × role) → `teams`/`team_members`.
- `ops_projects` (optionally linked to the public `projects` row by `public_project_id`) → `milestones`, `project_members`.
- `work_orders` with per-organisation numbering (`work_order_counters`, `SELECT … FOR UPDATE` + retry), `work_order_assignees` (user or team), `work_order_watchers`, `work_order_categories`, `work_order_assets`, `work_order_status_history`, `time_entries`, `cost_entries`. Recurring work orders carry `recurring_rule_json`; each generated occurrence has a unique `generation_key` so re-completion cannot double-generate.
- `assets` (tree via `parent_asset_id`) with `asset_status_history` and `asset_types`.
- `comments`, `attachments`, `notifications`, `saved_filters`, `ops_audit_events` are polymorphic on `(entity_type, entity_id)`.

Legacy tables (`users`, `projects`, `members`, inventory, prints, calendar …) are untouched. Legacy roles map to ops roles once (`admin→chapter_admin`, `team_leader→team_lead`, `member→full_member`) and the legacy `users.role` column is kept in sync when an admin changes a member's ops role.

## Frontend layout (`apps/ops-web/src`)

| Path | Responsibility |
| --- | --- |
| `api/client.ts` | `api()` fetch wrapper: CSRF header, envelope unwrapping, typed `ApiError` with field errors, 401 listener. |
| `api/hooks.ts` | TanStack Query hooks and cache keys; `useApiMutation` invalidates dependent queries. |
| `contracts/` | TypeScript mirror of the pydantic models (`types.ts`) and zod form schemas (`schemas.ts`). |
| `ui/` | Design tokens (`tokens.css`), primitives (Button, Badge, Avatar, Field, Picker, FilterChip, Popover/DropdownMenu, Dialog, SideSheet, Tabs, Toast, EmptyState/Skeleton/ErrorState, ActivityTimeline, Comments, Attachments, ReportCard). |
| `shell/AppShell.tsx` | Sidebar with Setup / Work / Optimize / Manage groups (built modules only, permission-filtered), setup banner, notifications, account menu, mobile drawer. |
| `auth/SessionProvider.tsx` | Session loading, `can()`, redirect to `/app/login`, `RequirePermission` guard. |
| `pages/` | Login, Setup Center, Reporting (operations dashboard), Work Orders (list + detail + create/edit pane), Projects (list, detail tabs), Teams & Users, Locations, Categories, Assets, Settings (chapter, roles, audit). |
| `lib/listParams.ts` | URL-backed list state (`filter[...]`, `sort`, `q`, extras) so views are shareable. |

Routing is client-side under the `/app` base; Flask serves `index.html` for every unknown `/app/*` path. Unknown API errors are surfaced inline (never a blank screen); a 401 anywhere returns the user to the login page with a `next` parameter.

## Cross-cutting

- **Tenancy**: a single default organisation (`asme-uiowa`) exists today; every query is scoped by `organization_id` so a second chapter can be added without schema changes.
- **Audit**: every create/update/status change writes an `ops_audit_events` row; the work-order history panel and project activity tab read from it.
- **Notifications**: in-app rows for assignment, status change, comments/mentions, watchers. Email delivery is not wired in this release.
- **Files**: stored under the instance folder, served through short-lived signed URLs; virus scanning is a stub (`scan_status = skipped`).
- **Feature flag**: `ASME_OPS_ENABLED` (default on) registers the ops API and SPA; turning it off leaves the legacy site exactly as it was.
