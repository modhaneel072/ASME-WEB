# ADR-0001 — Transitional modular monolith: Flask API + React/TypeScript Ops frontend

**Status:** Accepted — 2026-09-08
**Deciders:** principal engineer (this build)
**Context:** ASME Ops master prompt §3 (target architecture), §32 (recommended NestJS/Next.js stack), §2 (preservation rules)

## Context

The repository already contains a working Flask application that serves the public
ASME website, the member/admin portal, inventory checkout/return, NFC attendance and
kiosk flows, the print queue, and room scheduling. Immediately before this build it
was restructured into the `asme/` package: an application factory, typed settings,
eight blueprints, a services layer that owns every transaction, adapter-based
integrations (Google/Outlook calendar, SMTP, Anthropic), a transactional outbox with an
in-process worker, Alembic migrations (`0001_baseline`, `0002_launchpad`) and 155
passing tests. The database is SQLite in development and PostgreSQL in production.

The master prompt prefers an `apps/` + `packages/` monorepo with a NestJS API and a
Next.js client, but explicitly allows a transitional architecture that keeps Flask and
PostgreSQL for existing routes while adding a versioned JSON API and a React/TypeScript
Ops frontend, and warns against rewriting working systems "merely for fashion".

## Decision

1. **Keep Flask (`asme/`) as the API, public site and worker process.** It already
   provides the capabilities the target stack must have: real migrations, a service
   layer, durable background jobs, audit logging, typed error envelopes and tests.
   Rewriting it in NestJS would endanger working behaviour for no user-visible gain.
2. **Add an organization-scoped ops domain in `asme/ops/`.** New tables use UUID
   primary keys, carry `organization_id`, `created_at/updated_at/created_by/updated_by`,
   and are added by additive, reversible Alembic migrations. Existing tables are never
   renamed or dropped; existing rows are mapped in (users → memberships, public
   projects → ops projects) by repeatable backfills.
3. **Add a React 18 + TypeScript Ops frontend in `apps/ops-web/`** (Vite, TanStack
   Query, React Hook Form + Zod, Recharts, lucide-react icons) with an original
   component system and design tokens taken from spec §6. Flask serves the built
   bundle under `/app/*`; in development Vite proxies `/api` to Flask.
4. **Contracts** are pydantic v2 models on every ops endpoint; an OpenAPI document is
   generated from them at `/api/v1/ops/openapi.json`; the frontend keeps mirrored
   TypeScript/Zod types in `apps/ops-web/src/contracts/`.
5. **Authorization** is permission-based (`work_order.create`, `request.review`, …)
   with scopes (chapter / project / team / assigned / own) resolved server-side per
   record. Legacy roles map to system roles (`member → Full Member`,
   `team_leader → Team Lead`, `admin → Chapter Administrator`).
6. **Feature flag:** `ASME_OPS_ENABLED` gates the ops blueprint and `/app` routes.

### Mapping to the preferred repository shape

| Preferred | Realised as | Why not moved |
|---|---|---|
| `apps/api`, `apps/worker` | `asme/` (blueprints, services, jobs) | moving 12k lines changes every import and every deploy config for no behaviour change |
| `apps/ops-web` | `apps/ops-web/` | as specified |
| `apps/public-web` | `templates/site`, `static/`, `asme/blueprints/public.py` | preserved in place |
| `packages/ui` | `apps/ops-web/src/ui/` | single consumer today |
| `packages/contracts` | `asme/ops/schemas/` + `apps/ops-web/src/contracts/` | pydantic is the source of truth |
| `packages/database` | `asme/models/`, `asme/ops/models/`, `migrations/` | one Alembic history |
| `packages/auth` | `asme/ops/authz.py` | |
| `packages/reporting` | `asme/ops/services/dashboard.py` + `docs/reporting-metrics.md` | |

## Alternatives considered

- **NestJS + Next.js rewrite (spec §32).** Rejected for V1: it duplicates a working
  backend, forces a second ORM/migration history against the same database, and the
  preservation rules (§2) make a big-bang cutover the riskiest option.
- **Server-rendered Jinja + HTMX ops UI.** Rejected: the spec's interaction model
  (right-side create panes that keep list context, saved views, live filters,
  master-detail with deep links) is markedly easier to get right and test as an SPA,
  and the spec asks for React/TypeScript.
- **Extend the legacy `projects`/`events` tables in place.** Partially rejected:
  operational projects get their own `ops_projects` table linked to the public
  `projects` row through `public_project_id`, because public content (slug, gallery,
  copy) and operational programme data (lead, milestones, budget, risk) have different
  owners and lifecycles.

## Consequences

- Two languages in the repository (Python + TypeScript). Accepted; the boundary is a
  versioned JSON API with generated OpenAPI.
- Realtime messaging (Stage 3) will need a WebSocket gateway; the interim fallback is
  query polling with revalidation, which the client already supports.
- Production deploys must build the frontend (`npm ci && npm run build` in
  `apps/ops-web`) before starting Flask; see `docs/deployment.md`.

## Migration impact and rollback

- Migration `0003_ops_foundation` is additive (new tables, one nullable column on
  `projects`). Downgrade drops only those tables/columns.
- Backfills are repeatable (`python manage.py upgrade` re-runs seeds idempotently).
- Rollback: set `ASME_OPS_ENABLED=0` (ops API and `/app` disappear; legacy portal is
  untouched), then `flask db downgrade 0002_launchpad` if the schema must go too.
