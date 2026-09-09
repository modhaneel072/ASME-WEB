# ASME @ UIowa Web Platform

One Flask application, packaged as `asme/`, serving:

- the public marketing site
- the member / team-leader / admin portal
- inventory checkout + returns with an append-only stock ledger
- the 3D print queue
- NFC attendance, kiosk login and meeting check-in
- room scheduling mirrored to Google Calendar or Outlook
- **Launchpad** – the onboarding engine that turns a new signup into someone trusted with the shop
- **ASME Ops** – work orders, projects, assets and teams for the chapter, served at `/app` (see below)

The design this codebase implements is the *ASME Backend Blueprint* (sections 00–09); the layout below follows it.

## ASME Ops (operations platform at `/app`)

ASME Ops is the chapter's work-management app: work orders, projects, assets, teams, locations,
categories and an operations dashboard, with a Setup Center that tracks what is actually configured.
It is a React 18 + TypeScript SPA in `apps/ops-web`, served by this Flask app under `/app` and backed
by the `/api/v1` ops API (`asme/ops`). The legacy portal moved to `/legacy/app`.

```powershell
cd apps/ops-web
npm ci
npm run build                 # -> apps/ops-web/dist, served by Flask at /app
cd ../..
python manage.py seed-demo    # development only: Crater Cruncher Rover demo chapter, password ChangeMe123!
python manage.py serve        # http://127.0.0.1:5000/app  (sign in as admin@uiowa.edu)
```

For frontend development run `npm run dev` in `apps/ops-web` and open `http://localhost:5173/app`
(the Vite dev server proxies `/api` to Flask on :5000).

Quality gates: `npm run typecheck`, `npm run lint`, `npm run test` (vitest), `npm run test:e2e`
(Playwright against a throwaway seeded server started by `python manage.py serve-e2e`), and `python -m pytest`.

Docs: [architecture](docs/architecture/overview.md) · [API](docs/api.md) · [permissions](docs/permissions-matrix.md) ·
[reporting metrics](docs/reporting-metrics.md) · [migration plan](docs/migration-plan.md) · [deployment](docs/deployment.md) ·
[test plan](docs/test-plan.md) · [implementation status](docs/implementation-status.md) · [ADR-0001](docs/decisions/ADR-0001-transitional-architecture.md).

## Layout

```
app.py                 WSGI entry point (gunicorn app:app) – 12 lines
manage.py              upgrade / seed / evaluate / reconcile / worker / serve / routes
asme/
  __init__.py          create_app() factory
  config.py            typed Settings – every ASME_* var resolved once, validated at boot
  constants.py         roles, printers, rooms, entitlement keys
  content_data.py      static site copy + exec profile overrides
  models/              ORM, split by domain (identity, inventory, fabrication, events, content, audit, onboarding)
  auth/                session, role + entitlement decorators, login rate limiter
  blueprints/          public · auth · kiosk · portal · admin · api_v1 · api_legacy · legacy_ops
  services/            the only code that opens a transaction
    onboarding/        Launchpad engine: rules · engine · entitlements · seeds
  integrations/        calendar (google/outlook adapters), mail, assistant (Anthropic)
  events/              in-process domain event bus (services emit, Launchpad subscribes)
  jobs/                transactional outbox + in-process worker + handlers
  ops/                 ASME Ops domain: models · authz · tenancy · services · api (/api/v1) · schemas · seeds · web (/app)
apps/ops-web/          ASME Ops frontend (React 18 + TypeScript + Vite); build output served by asme/ops/web.py
migrations/            Alembic (Flask-Migrate) – 0001_baseline, 0002_launchpad, 0003_ops_foundation
tests/                 pytest suite (app factory, in-memory SQLite)
templates/ static/     unchanged; two new pages: portal/member_launchpad.html, portal/admin_launchpad.html
models.py              compatibility shim -> asme.models
services/ai_assistant.py  compatibility shim -> asme.integrations.assistant
```

**Layering rule.** Blueprints parse, authorize and render. Services own transactions and business rules and are the only place that touches `db.session`. Integrations are adapters behind an interface. A web form, a JSON client and a kiosk tap all end up in the same `inventory.checkout()`.

## Quick start (Windows PowerShell)

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements-dev.txt
Copy-Item .env.example .env       # sets ASME_ENV=development
python manage.py upgrade          # migrate to head + seed defaults (admin, projects, Launchpad tracks)
python db_init.py                 # optional: sample members/items/events for local testing
python app.py
```

Open `http://127.0.0.1:5000`. Default admin is `ASME_DEFAULT_ADMIN_EMAIL` / `ASME_DEFAULT_ADMIN_PASSWORD` from `.env`.

Run the tests:

```powershell
python -m pytest
```

## Environments

`ASME_ENV` is `development`, `production` or `testing`.

| Setting | development | production |
|---|---|---|
| `ASME_AUTO_MIGRATE` (migrate + seed at boot) | on | **off** – run `python manage.py upgrade` as a release step |
| template auto-reload | on | off |
| default `ASME_SECRET_KEY` | warning | **refuses to start** |
| `ASME_ONBOARDING_ENFORCE` | shadow (0) | shadow (0) until you flip it |

Production deploys already run migrations for you:

- **Procfile** – `release: python manage.py upgrade`
- **render.yaml** – `startCommand: python manage.py upgrade && gunicorn …`
- **Elastic Beanstalk** – `.ebextensions/03_migrate.config` (`leader_only`)

### Existing databases

A database created by the old `db.create_all()` code has no `alembic_version` table. `manage.py upgrade` detects that, stamps it at `0001_baseline` (the pre-package schema) and upgrades to head. Migration `0002` also backfills `transactions.user_id` from the legacy member's email, generates missing inventory tracking ids, and copies legacy `items.nfc_tag` values into `item_tags`. The old runtime `ALTER TABLE` code is gone.

## Launchpad (onboarding engine)

Two tracks share one engine. Task state is keyed on `(subject_type, subject_id)`, so the same tables drive both.

**Member track** – `signed_up` → `shop_ready` → `contributor` → `team_lead`. Finishing a phase grants entitlements:

| Phase | Grants |
|---|---|
| Signed up | `portal_access`, `event_checkin` |
| Shop ready (safety module ≥ 80 %, one meeting, one supervised checkout) | `shop_access`, `print_submit` |
| Contributor (3 meetings, 3 clean returns, 10 logged hours) | `extended_loan_14d`, `room_booking` |
| Team lead (nominated, lead training, build cycle signed off) | `role:team_leader`, `print_approve`, `checkout_signoff` |

**Chapter track** – `stand_up` (roster ≥ 20, teams ≥ 3, items ≥ 40, locations ≥ 5, calendar connected, first NFC card) → `run_semester` (schedule published, print queue open, loan periods set, 60 % of members shop-ready). Subject is the academic year, e.g. `chapter:2026-27`.

Progress is derived, never stored as a number: every domain event (checkout, return, check-in, training, team join, hours logged…) re-runs the rule evaluators against the tables that already hold the truth. If a requirement stops being met (training expired, a tool goes overdue) the phase regresses and its grants are revoked.

**Enforcement.** Routes are gated with `@require_entitlement("shop_access")` etc. With `ASME_ONBOARDING_ENFORCE=0` (default) the gate runs in *shadow mode*: it logs `shadow_blocks` on the request line and lets the request through. Watch the logs for a few weeks, then set `ASME_ONBOARDING_ENFORCE=1`. Admins always pass; admins can also grant/revoke override entitlements (with expiry) from **Admin → Launchpad**.

Pages: `/portal/member/launchpad`, `/portal/leader/launchpad` (sign-offs, training), `/portal/admin/launchpad` (chapter track, overrides, re-evaluate).

## API v1

Versioned JSON under `/api/v1`. One error shape everywhere: `{"ok": false, "code": "...", "error": "..."}`; a missing entitlement returns `403` with `entitlement` and the `phase` that unlocks it. Collections use cursor pagination; `/items` sends an `ETag`.

| Method | Path | Gate |
|---|---|---|
| GET | `/api/v1/me` | login |
| GET | `/api/v1/onboarding` (`?user_id=` for leads, `&include_chapter=1` for admins) | login |
| POST | `/api/v1/onboarding/tasks/:key/complete` | manual/sign-off tasks only |
| POST | `/api/v1/onboarding/training` | team_leader |
| GET | `/api/v1/items`, `/api/v1/items/by-tag/:uid` | `portal_access` |
| POST | `/api/v1/checkouts` (requires `Idempotency-Key`) | `shop_access` |
| POST | `/api/v1/checkouts/:id/return` | `shop_access` |
| GET | `/api/v1/checkouts?state=open` | `portal_access` |
| GET/POST | `/api/v1/print-requests`, PATCH `/api/v1/print-requests/:id` | `print_submit` / `print_approve` |
| POST | `/api/v1/checkins` (resolves the event server-side; opens an `open_shop` event if none) | `event_checkin` |
| GET | `/api/v1/availability`, POST `/api/v1/bookings` | `room_booking` |
| POST | `/api/v1/roster/import` | admin |

The old `/api/*` kiosk endpoints are unchanged in shape and now call the same services.

## Background work

`outbox_jobs` is a transactional outbox: a service commits the job row together with the domain write, and an in-process worker thread (started on the first request; `ASME_OUTBOX_WORKER=0` to disable, or run `python manage.py worker` as a separate process) executes it with retry/backoff. Handlers: `calendar.create_event`, `calendar.delete_event`, `mail.send`, `stock.reconcile` (nightly; files `stock_discrepancies` for admin review instead of silently correcting), `onboarding.evaluate`, `onboarding.evaluate_all`. Failed jobs and discrepancies show on **Admin → Settings** and **Admin → Inventory** with retry/resolve buttons.

## Roles

`member < team_leader < admin` (`ROLE_ORDER`). Roles say what job you hold; entitlements say what you have proven. Both are checked.

## Key routes

Public: `/`, `/who-we-are`, `/executive-team`, `/projects`, `/projects/<slug>`, `/events`, `/gallery`, `/join`, `/contact`, `/sponsors`, `/login`, `/signup`, `/forgot-password`, `/admin-login`, `/healthz`.

Portal: `/portal` (router), `/portal/member`, `/portal/member/{inventory,my-items,prints,calendar,profile,help,launchpad}`, `/portal/member/schedule` (team_leader + `room_booking`), `/portal/leader/launchpad`.

Admin: `/portal/admin`, `/portal/admin/{members,attendance,inventory,prints,calendar,settings,launchpad}` plus the POST actions under each.

Shared-device: `/kiosk`, `/checkin`, `/checkin/select`, `/checkin/success`, `/pair/member`, `/pair/item`.

Legacy ops pages stay behind `ASME_ENABLE_LEGACY_OPS=1`; otherwise `/dashboard`, `/attendance`, `/inventory`, `/prints`, `/activity`, `/settings`, `/scan`, `/my-items`, `/admin/nfc`, `/calendar`, `/app` redirect into the portal.

`python manage.py routes` prints every URL rule with its endpoint.

## Calendar

`ASME_CALENDAR_PROVIDER=google|outlook` selects one adapter in `asme/integrations/calendar/`; `services/scheduling.py` never learns which. Google needs `GOOGLE_SERVICE_ACCOUNT_JSON`, `GOOGLE_CALENDAR_ID_ROBOTICS`, `GOOGLE_CALENDAR_ID_FLUIDS`; Outlook needs the `ASME_OUTLOOK_*` variables. Bookings commit an `Event` + `CalendarSync(pending)` + outbox job; the link appears once the worker syncs it. `python scripts/outlook_sync_doctor.py` still diagnoses Outlook.

## Environment variables

See `.env.example`. New since the rewrite: `ASME_ENV`, `ASME_AUTO_MIGRATE`, `ASME_ONBOARDING_ENFORCE`, `ASME_OUTBOX_WORKER`, `ASME_OUTBOX_POLL_SECONDS`, `ANTHROPIC_API_KEY`, `ASME_ASSISTANT_MODEL`. Everything the old app read is still honoured.

## Bulk roster import

Admin → Members → **Bulk Import Roster**: upload the roster PDF; accounts are created/updated idempotently (username = first initial + last name, password = 2 letters + initial + 5 digits) and a credentials CSV downloads. Also available as `POST /api/v1/roster/import`.

## What is deliberately still here

- `members` table and `member_id` columns – `User` is canonical and every service writes `user_id`; the legacy table stays one release as the compatibility layer for the kiosk flows. Dropping it is a separate migration after the backfill has been verified on production data.
- `PrintJob`, `AttendanceScan`, `Meeting` – legacy-ops models, read/written only when `ASME_ENABLE_LEGACY_OPS=1`.
