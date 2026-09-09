# Migration plan: legacy portal → ASME Ops

## Principles (from the master prompt, enforced in code and tests)

- No destructive schema commands against an existing database. All changes are additive Alembic migrations (`migrations/versions/0003_ops_foundation_*`), applied with `python manage.py upgrade` or automatically at boot outside production (`ASME_AUTO_MIGRATE`).
- Existing password hashes are never rewritten; ASME Ops authenticates through the same `users` table and `identity.authenticate`.
- Populated legacy tables and columns are never renamed or dropped. `members` stays for one more release (see `asme-backend-rewrite` notes), `projects` remains the public-site source of truth and is linked from `ops_projects.public_project_id`.
- Secrets are never committed. `.env` is git-ignored; `.env.example` documents every variable.

## Stages

| Stage | State | What happens |
| --- | --- | --- |
| 0 – Baseline | done | Docs and ADR written, existing 155 backend tests green, no product change. |
| 1 – Foundation | done | Migration 0003 adds `organizations`, `memberships`, roles/permissions, teams, locations, categories, assets, comments, attachments, notifications, saved filters, audit events, work-order and project tables. `seed_defaults()` (run at boot / `manage.py upgrade`) creates the default organisation, the 12 system roles with their permission grants, default categories and asset types, and a membership for every active user with the mapped legacy role. Idempotent: re-running adds nothing. |
| 2 – Projects + Work Orders slice | done | React app under `/app`, Setup Center, projects, work orders, teams/users, locations, categories, assets, reporting. Legacy portal continues to work at `/legacy/app` (when `ASME_ENABLE_LEGACY_OPS=1`) and the public site is untouched. |
| 3 – Requests, parts, procedures | not started | Listed in the Setup Center as "Not in this release"; no navigation entries exist. |
| 4 – Maintenance plans, automations, dashboards | not started | Same. |
| 5 – Retire legacy portal | not started | Only after every legacy screen has an ops equivalent and members have moved. Requires a separate ADR. |

## Data migration details (stage 1, already applied)

| Legacy | ASME Ops | Rule |
| --- | --- | --- |
| `users.role` (`admin`, `team_leader`, `member`) | `memberships.role_id` | `LEGACY_ROLE_MAP`; changing a member's ops role writes the reverse mapping back to `users.role` so the legacy portal keeps working. |
| `users` (all active) | `memberships` (status `active`) | one membership per user per organisation; inactive users get no membership until reactivated. |
| `projects` (public site) | `ops_projects.public_project_id` | optional link set by the demo seed for the rover; new ops projects may set it from the project form later. |
| nothing | `locations` "Default" | created per organisation; assets default to it. |

Rollback: migration 0003 has a full `downgrade()`, but because it only adds tables the safe rollback in production is to set `ASME_OPS_ENABLED=0`, which unregisters the API and SPA without touching data.

## Cut-over checklist for a real deployment

1. Take a database backup (`pg_dump` on Render/EB).
2. Deploy the build; `python manage.py upgrade` runs migration 0003 and the idempotent seed.
3. Verify `GET /api/v1/ops/health` and `GET /app` render.
4. Sign in as an existing admin; complete the Setup Center foundation phase.
5. Change `ASME_DEFAULT_ADMIN_PASSWORD` if it is still a default (startup warning in production).
6. Do **not** run `manage.py seed-demo` against production; it creates development accounts with a documented development-only password.
