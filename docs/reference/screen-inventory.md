# Screen inventory

Status legend: **BUILT** (persisted, authorised, tested), **PARTIAL** (some states or
actions missing — listed), **PLANNED** (not started; no navigation item is shown for it
so nothing decorative appears in the product). This table matches the code as of this commit; the
earlier draft of this file described intent and has been reconciled with what shipped.

| Route (under `/app`) | Screen | Status | Notes |
|---|---|---|---|
| `/login` | Ops sign-in (email or username) | BUILT | reuses the existing password auth + rate limiter; `/auth/login` redirects here |
| `/setup` | Setup Center | BUILT | 3 phases from spec §8; completion computed from persisted rows; tasks for unbuilt modules show "Not in this release" with no CTA; officer guide tab; banner hide/show |
| `/reporting` | Operations dashboard | BUILT | real aggregates, 7/30/90/365-day range, project filter; formulas in `docs/reporting-metrics.md` |
| `/projects` | Project cards | BUILT | Active / All / Archived, status/risk/lead filters, search, sort; create pane |
| `/projects/:id` (`overview`, `work-orders`, `milestones`, `members`, `files`, `activity`) | Project detail tabs | BUILT | health KPIs, teams, milestones CRUD, members dialog (project role + team), files, audit activity; edit pane; archive. Budget lives on Overview (amount, code, used); there is no separate Budget or Assets tab |
| `/work-orders` | Work-order list | BUILT | To Do / Done tabs with counts, 11 filter chips, search (title/description/#number), 7 sorts, saved views (private/team/chapter), panel and table layouts, cursor "Load more"; Calendar and Workload layouts are PLANNED and not offered |
| `/work-orders?new=1` | Right-side create pane | BUILT | title, description, priority (Critical gated by permission), type, categories, project, team, location, asset + related assets, people/team assignees, watchers, start/due, estimate, budget code, recurrence, save as draft, create-another |
| `/work-orders?wo=:id` (and `/work-orders/:id` deep link) | Work-order detail | BUILT | status actions by transition + permission (Open/Start/Hold/Resume/Complete/Cancel), assign, watchers, blocked flag, edit pane, duplicate, copy link, sub-work orders with progress, comments (edit/delete own), files (signed links), time & cost entries, status history + full audit trail |
| `/teams-users` (`/users`) | Teams and users | BUILT | teams table + create/edit pane with leads; users table with inline role/status change (last-admin guard on the server), invite pane |
| `/locations` | Locations tree | BUILT | nested rows, default location, archive, create/edit pane |
| `/categories` | Categories | BUILT | colour + icon, usage counts, archive, duplicate-name error surfaced |
| `/assets` | Assets (table + detail panel) | BUILT | status/criticality/project/location/team/type filters; detail with sub-assets, files, status history; create/edit pane; status change with downtime type/reason. Meters and custom-field editor are PLANNED |
| `/settings` (`chapter`, `roles`, `audit`) | Settings | BUILT | chapter profile (org.manage), read-only role/permission matrix, audit log with entity/actor filters |
| `/requests`, `/messages`, `/events` | | PLANNED | Stage 3 – not in navigation |
| `/parts`, `/purchase-requests`, `/vendors`, `/sponsors` | | PLANNED | Stage 4 |
| `/library/*`, `/maintenance-plans` | | PLANNED | Stage 5 |
| `/meters`, `/automations` | | PLANNED | Stage 6 |
| `/reporting/*` beyond the operations dashboard | | PLANNED | Stage 7 |
| notifications / integrations settings | | PLANNED | in-app notification list exists in the shell; email delivery and integrations are not wired |

## Responsive behaviour (verified by `e2e/responsive.spec.ts`)

| Viewport | Sidebar | Work-order list + detail | Create pane |
|---|---|---|---|
| 1440×900, 1366×768 | fixed 264px | two columns (list 5fr / detail 6fr), detail sticky | 560px right sheet over the list |
| 768×1024 | drawer behind a top bar | detail replaces the list with a back button | 560px right sheet |
| 390×844 | drawer | same as tablet | full-width sheet |
