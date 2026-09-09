# Reporting metrics

Every number on the Operations dashboard (`GET /api/v1/ops/dashboard/operations`, `apps/ops-web/src/pages/DashboardPage.tsx`) and project health (`GET /api/v1/projects/{id}/health`) is computed live from persisted rows by `asme/ops/services/dashboard.py` and `asme/ops/services/projects.py`. Nothing is cached or pre-aggregated. The formulas below are the contract; `tests/ops/test_dashboard.py` pins them.

## Range

`range` is one of `7d`, `30d` (default), `90d`, `365d`. `start = now − days`, `end = now` (UTC). Optional `project_id` restricts every metric to work orders of that project.

## Operations dashboard

| Metric | Definition |
| --- | --- |
| Open | count of work orders with status in `{DRAFT, OPEN, IN_PROGRESS, ON_HOLD}` (not range-limited) |
| Overdue | open work orders with `due_at < now` |
| Blocked | open work orders with `is_blocked = true` |
| Due in 7 days | open work orders with `now ≤ due_at ≤ now + 7d` |
| Created | work orders with `created_at` in range |
| Completed | work orders with status `DONE` and `completed_at` in range |
| Created vs completed | weekly buckets from `start` (`days // 7` buckets, the last one stretched to `end` so no day is dropped); each bucket counts `created_at` / `completed_at` in `[bucket_start, bucket_end)` |
| On-time completion rate | among work orders completed in range: `completed_at ≤ due_at` or no due date → on time; rate = on_time / completed × 100 (rounded). `null` when nothing was completed |
| Average time to complete | mean of `(completed_at − created_at)` in hours over work orders completed in range, 1 decimal; `null` when none |
| By status | count per status over all work orders (not range-limited) |
| By priority | open work orders per priority |
| Created by work type | work orders created in range per `work_type` |
| Repeating vs one-off | work orders created in range with / without `recurring_rule_json` |
| Workload by team | open work orders grouped by `team_id`, top 10 |
| Workload by person | open work orders grouped by assignee user, top 10 (a work order with two assignees counts once per person) |
| Hours logged | sum of `time_entries.minutes` created in range ÷ 60, 1 decimal |
| Costs | sum of `cost_entries.amount` created in range, grouped by `type` (`part`, `labor`, `vendor`, `other`) and total |

## Project health

| Metric | Definition |
| --- | --- |
| Completion % | if the project has milestones: `Σ weight(complete milestones) / Σ weight(all milestones) × 100` (weights default to 1); without milestones: `done_wo / total_wo × 100`; 0 when there is nothing to measure. Implemented in `projects.completion_percent`. |
| Open / done / total / overdue / blocked | work-order counts for the project (`overdue` and `blocked` restricted to open statuses) |
| By status / by priority | counts of the project's work orders (priority restricted to open) |
| Workload by team | open work orders per team |
| Milestones | total, complete, `at_risk` = overdue or due within 14 days and not complete, `next` = earliest incomplete milestone with a due date (`days_left` may be negative) |
| Budget | `amount` from the project, `used` = Σ cost entries on the project's work orders, `remaining = amount − used` (null without an amount) |
| Hours logged | Σ `work_orders.actual_minutes` ÷ 60 |

## Work-order list counts

The `counts` object on `GET /work-orders` is computed on the visible set for the caller (`todo` = open statuses, `done` = `{DONE, CANCELED, SKIPPED}`), optionally restricted to `filter[project]`.

## Known limitations

- All aggregates are computed per request with SQL `GROUP BY`; no materialised views. Fine for chapter-scale data (thousands of rows), revisit if a chapter exceeds ~100k work orders.
- Timezone: buckets and "today" use UTC on the server; the browser renders dates in the viewer's local zone. The organisation's timezone setting is stored but not yet applied to bucket boundaries.
