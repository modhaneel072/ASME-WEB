# Permissions matrix

Authorisation is permission-based. A **role** is a named bundle of `(permission, scope)`
pairs; a **membership** binds a user to an organisation with one role. Every private
query is filtered by `organization_id`; every mutation loads the record, resolves its
organisation/project/team associations, then checks the permission at the required
scope. Hiding a control in the UI is never the authorisation.

Source of truth: `asme/ops/authz.py` (`PERMISSIONS`, `SYSTEM_ROLES`, `LEGACY_ROLE_MAP`).
Tests: `tests/ops/test_authz.py` (allowed + denied + ID-tampering cases).

## Scopes

| Scope | Meaning |
|---|---|
| `chapter` | any record in the organisation |
| `project` | records whose project the user leads or belongs to |
| `team` | records whose team (or assignee team) the user leads |
| `assigned` | records where the user is an assignee (directly or via an assignee team they belong to) |
| `own` | records the user created |

A permission held at a wider scope satisfies a narrower requirement. Chapter
Administrator holds every permission at `chapter` scope.

## Legacy role mapping

| Legacy `users.role` | System role key | Applied when |
|---|---|---|
| `admin` | `chapter_admin` | migration `0003` backfill + first ops sign-in |
| `team_leader` | `team_lead` | same |
| `member` | `full_member` | same |

Changing a role in ASME Ops writes the reverse mapping back to `users.role`
(`chapter_admin`/`executive_officer` → `admin`; `project_lead`/`team_lead` →
`team_leader`; everything else → `member`) so the legacy portal stays consistent.

## Matrix

✔ = chapter scope · P = project scope · T = team scope · A = assigned only · O = own only · — = not granted

| Permission | Chapter Admin | Exec Officer | Project Lead | Team Lead | Full Member | Shop Operator | Requester | Inventory Mgr | Safety Officer | Treasurer | Faculty Advisor | Sponsor/Guest |
|---|---|---|---|---|---|---|---|---|---|---|---|---|
| `org.read` | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ |
| `org.manage` | ✔ | — | — | — | — | — | — | — | — | — | — | — |
| `setup.manage` | ✔ | ✔ | — | — | — | — | — | — | — | — | — | — |
| `user.manage` | ✔ | ✔ | — | — | — | — | — | — | — | — | — | — |
| `role.manage` | ✔ | — | — | — | — | — | — | — | — | — | — | — |
| `team.read` | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | — |
| `team.manage` | ✔ | ✔ | P | T | — | — | — | — | — | — | — | — |
| `location.manage` | ✔ | ✔ | — | — | — | — | — | ✔ | — | — | — | — |
| `category.manage` | ✔ | ✔ | — | — | — | — | — | — | — | — | — | — |
| `asset.read` | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | — | ✔ | ✔ | ✔ | ✔ | — |
| `asset.manage` | ✔ | ✔ | P | — | — | — | — | ✔ | — | — | — | — |
| `project.read` | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | — |
| `project.create` | ✔ | ✔ | — | — | — | — | — | — | — | — | — | — |
| `project.manage` | ✔ | ✔ | P | — | — | — | — | — | — | — | — | — |
| `work_order.read_all` | ✔ | ✔ | ✔ | ✔ | ✔ | — | — | ✔ | ✔ | ✔ | ✔ | — |
| `work_order.read_assigned` | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | — | ✔ | ✔ | ✔ | ✔ | — |
| `work_order.create` | ✔ | ✔ | P | T | O | — | — | ✔ | ✔ | — | — | — |
| `work_order.edit` | ✔ | ✔ | P | T | O | — | — | ✔ | ✔ | — | — | — |
| `work_order.assign` | ✔ | ✔ | P | T | — | — | — | ✔ | ✔ | — | — | — |
| `work_order.start` | ✔ | ✔ | P | T | A | A | — | ✔ | ✔ | — | — | — |
| `work_order.complete` | ✔ | ✔ | P | T | A | A | — | ✔ | ✔ | — | — | — |
| `work_order.cancel` | ✔ | ✔ | P | T | O | — | — | ✔ | ✔ | — | — | — |
| `work_order.set_critical` | ✔ | ✔ | P | — | — | — | — | — | ✔ | — | — | — |
| `comment.create` | ✔ | ✔ | ✔ | ✔ | ✔ | A | — | ✔ | ✔ | ✔ | ✔ | — |
| `file.upload` | ✔ | ✔ | ✔ | ✔ | ✔ | A | — | ✔ | ✔ | ✔ | — | — |
| `time_entry.create` | ✔ | ✔ | ✔ | ✔ | A | A | — | ✔ | ✔ | — | — | — |
| `cost_entry.create` | ✔ | ✔ | P | T | — | — | — | ✔ | — | ✔ | — | — |
| `request.submit` | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | — |
| `request.review` | ✔ | ✔ | P | T | — | — | — | ✔ | ✔ | — | — | — |
| `inventory.read` | ✔ | ✔ | ✔ | ✔ | ✔ | ✔ | — | ✔ | ✔ | ✔ | ✔ | — |
| `inventory.manage` | ✔ | — | — | — | — | — | — | ✔ | — | — | — | — |
| `purchase.review` | ✔ | ✔ | — | — | — | — | — | — | — | ✔ | ✔ | — |
| `procedure.publish` | ✔ | ✔ | — | — | — | — | — | — | ✔ | — | — | — |
| `safety.manage` | ✔ | — | — | — | — | — | — | — | ✔ | — | — | — |
| `report.view` | ✔ | ✔ | ✔ | ✔ | ✔ | — | — | ✔ | ✔ | ✔ | ✔ | shared only |
| `report.export` | ✔ | ✔ | P | — | — | — | — | ✔ | ✔ | ✔ | ✔ | — |
| `dashboard.manage` | ✔ | ✔ | ✔ | ✔ | — | — | — | — | — | — | — | — |
| `automation.manage` | ✔ | ✔ | — | — | — | — | — | — | — | — | — | — |
| `saved_filter.share` | ✔ | ✔ | ✔ | ✔ | — | — | — | ✔ | ✔ | — | — | — |
| `audit.read` | ✔ | ✔ | — | — | — | — | — | — | ✔ | ✔ | ✔ | — |

Permissions for modules not yet built (`request.*`, `inventory.*`, `purchase.*`,
`procedure.*`, `automation.*`) are defined now so roles are complete; the routes that
check them arrive with their stages.
