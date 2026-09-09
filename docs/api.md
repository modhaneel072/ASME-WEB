# ASME Ops API (v1)

Base path: `/api/v1`. Machine-readable contract: `GET /api/v1/ops/openapi.json` (OpenAPI 3.1, generated from the pydantic models). Everything below is implemented and covered by `tests/ops/`.

## Conventions

- **Auth**: signed session cookie (same login as the public site). `POST /ops/session/login` with `{identifier, password}` (email or username). Anonymous calls get `401 {ok:false, code:"login_required"}`.
- **CSRF**: every `POST/PUT/PATCH/DELETE` must send `X-Requested-With: ASME-Ops`; otherwise `403 code:"csrf"`.
- **Envelope**: `{ "ok": true, "payload": … }`. Errors: `{ "ok": false, "code": "<machine code>", "error": "<human message>", "details": [{"field","message"}]?, "permission"? }`.
- **Codes**: `validation` (400), `login_required` (401), `csrf`/`forbidden`/`link_expired` (403), `not_found` (404), `conflict`/`invalid_transition`/`name_taken` (409), `rate_limited` (429), `file_size` (413), `server_error` (500).
- **Lists**: `?filter[key]=a,b&sort=<key>&q=<text>&page[limit]=50&page[cursor]=<opaque>`. Unknown filter or sort keys are rejected with `validation`. Responses: `{items, next_cursor}`.
- **Tenancy**: ids from another organisation return 404, never 403.
- **Datetimes**: ISO-8601. Clients send offsets (`…Z`); the server stores and returns naive UTC.

## Endpoints

### Session, organisation, setup
| Method | Path | Permission | Notes |
| --- | --- | --- | --- |
| POST | `/ops/session/login` | public | rate limited per IP+identifier |
| POST | `/ops/session/logout` | session | |
| GET | `/ops/session` | session | user, organisation, role, permissions, grants, unread count |
| GET | `/ops/health` | public | |
| GET | `/ops/lookups` | session | users, teams, locations, categories, projects, assets, roles for selectors |
| GET/PATCH | `/org` | `org.read` / `org.manage` | PATCH also confirms the Setup Center profile task |
| GET | `/roles` | session | role → permission/scope matrix |
| GET | `/users` | `team.read` | `?include_inactive=1` needs `user.manage` |
| POST | `/users` | `user.manage` | invite / attach an existing account |
| PATCH | `/users/{id}` | `user.manage` | role_key, member_status, title; last-admin guard |
| GET | `/ops/setup` | session | phases/tasks computed from persisted rows |
| POST | `/ops/setup/banner` | session | `{dismissed}` per member |
| POST | `/ops/setup/guide-read` | session | |
| POST | `/ops/setup/confirm-profile` | `setup.manage` | |
| GET | `/ops/dashboard/operations` | `report.view` | `?range=7d|30d|90d|365d&project_id=` |
| GET | `/notifications` | session | `?unread=1` |
| POST | `/notifications/read-all`, `/notifications/{id}/read` | session | |
| GET | `/ops/audit` | `audit.read` | `?entity_type&entity_id&event_type&actor_id&offset&page[limit]` |
| GET | `/ops/openapi.json` | public | |

### Directory
| Method | Path | Permission |
| --- | --- | --- |
| GET | `/teams`, `/teams/{id}` | `team.read` |
| POST | `/teams` | `team.manage` (scope checked against `project_id`) |
| PATCH | `/teams/{id}` · PUT `/teams/{id}/members` · DELETE `/teams/{id}` (archive) | `team.manage` on the team |
| GET | `/locations`, `/locations/{id}` | session |
| POST/PATCH/DELETE | `/locations…` | `location.manage` (`is_default` via PATCH) |
| GET | `/categories`, `/categories/{id}` | session |
| POST/PATCH/DELETE | `/categories…` | `category.manage` |

### Projects
| Method | Path | Permission |
| --- | --- | --- |
| GET | `/projects` (`filter[lead,status,risk,year,team]`, `sort=name_asc|updated_desc|target_asc|risk_desc`, `view=active|all|archived`) | `project.read` |
| POST | `/projects` | `project.create` |
| GET | `/projects/{id}` (+`permissions`, `files`) · `/health` · `/activity` | `project.read` |
| PATCH | `/projects/{id}` · POST `/archive` · PUT `/members` | `project.manage` on the project |
| POST/PATCH/DELETE | `/projects/{id}/milestones[/{mid}]` | `project.manage` |

### Work orders
| Method | Path | Permission |
| --- | --- | --- |
| GET | `/work-orders` (`tab=todo|done`, `filter[status,priority,work_type,project,team,location,asset,category,assignee(me|unassigned|id),due(overdue|today|week|month|none|has),blocked,parent,created_by]`, `sort=due_asc|priority_desc|updated_desc|created_desc|number_desc|number_asc|project`) | `work_order.read_assigned`; visibility widened by `read_all` scope |
| POST | `/work-orders` | `work_order.create` (scope of project/team) |
| GET | `/work-orders/{id}` (+`comments`, `files`) | visible to assignee/creator/team/project/chapter per scope |
| PATCH | `/work-orders/{id}` | `work_order.edit` |
| POST | `/work-orders/{id}/start` · `/hold` · `/resume` | `work_order.start` |
| POST | `/work-orders/{id}/open` | `work_order.edit` (draft → open) |
| POST | `/work-orders/{id}/complete` (`completion_note, time_minutes, costs[], asset_status, follow_up_title`) | `work_order.complete` |
| POST | `/work-orders/{id}/cancel` | `work_order.cancel` |
| POST | `/work-orders/{id}/duplicate` · `/sub-work-orders` | `work_order.create` |
| PUT | `/work-orders/{id}/assignees` (`user_ids, team_ids`) | `work_order.assign` |
| PUT | `/work-orders/{id}/watchers` | `work_order.edit` |
| GET | `/work-orders/{id}/history` | visible |
| GET/POST | `/work-orders/{id}/comments` · PATCH/DELETE `/comments/{id}` | `comment.create`; edit/delete own |
| GET/POST | `/work-orders/{id}/time-entries` · `/cost-entries` | `time_entry.create` / `cost_entry.create` (logging for someone else needs `work_order.assign`) |

### Assets, files, saved filters
| Method | Path | Permission |
| --- | --- | --- |
| GET | `/assets` (`filter[status,criticality,project,location,team,parent,type]`), `/assets/{id}` (+children, files, history), `/asset-types` | `asset.read` |
| POST/PATCH | `/assets…` · POST `/assets/{id}/status` | `asset.manage` |
| GET | `/files?entity_type&entity_id` · `/files/{id}` | visible entity |
| POST | `/files` (multipart `file`, `entity_type`, `entity_id`) | `file.upload` on the entity |
| GET | `/files/{id}/download?token=` | signed token, expires after `ASME_FILE_URL_TTL_SECONDS` |
| DELETE | `/files/{id}` | uploader or entity editor |
| GET/POST | `/saved-filters?entity_type=` | session; `team`/`chapter` visibility needs `saved_filter.share` |
| PATCH/DELETE | `/saved-filters/{id}` | owner |

## Work-order state machine

```
DRAFT ──open──► OPEN ──start──► IN_PROGRESS ──complete──► DONE
                 │                │  ▲
                 │                hold  resume
                 │                ▼  │
                 └──cancel──► CANCELED ◄──cancel── ON_HOLD
```
`SKIPPED` is reserved for generated occurrences that are superseded. Invalid transitions return `409 code:"invalid_transition"` with the allowed targets. Completing a work order with `recurrence` creates the next occurrence exactly once; `follow_up_title` creates a linked follow-up; `asset_status` updates the primary asset and its status history.
