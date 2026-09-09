import { useState } from 'react';
import { Link } from 'react-router-dom';
import {
  Bar,
  BarChart,
  CartesianGrid,
  Legend,
  ResponsiveContainer,
  Tooltip as ChartTooltip,
  XAxis,
  YAxis,
} from 'recharts';
import { useDashboard, useLookups } from '@/api/hooks';
import { fmtMinutes, fmtMoney, humanize, STATUS_LABEL } from '@/lib/format';
import { projectFilterOptions } from '@/lib/options';
import { ErrorState, FilterChip, PageHeader, Segmented, Skeleton } from '@/ui';
import { BarList, KpiCard, ReportCard } from '@/ui/ReportCard';

const RANGES = [
  { value: '7d', label: '7 days' },
  { value: '30d', label: '30 days' },
  { value: '90d', label: '90 days' },
  { value: '365d', label: 'Year' },
];
const PRIORITY_COLORS: Record<string, string> = {
  NONE: '#c3ceda',
  LOW: '#7fb8e8',
  MEDIUM: '#ffcd00',
  HIGH: '#e58a00',
  CRITICAL: '#d84a4a',
};
const STATUS_COLORS: Record<string, string> = {
  DRAFT: '#c3ceda',
  OPEN: '#0878d1',
  IN_PROGRESS: '#e58a00',
  ON_HOLD: '#7c5ce7',
  DONE: '#00a878',
  CANCELED: '#d84a4a',
  SKIPPED: '#b91c1c',
};

/** Operations dashboard: formulas are documented in docs/reporting-metrics.md. */
export function DashboardPage() {
  const [range, setRange] = useState('30d');
  const [project, setProject] = useState<string[]>([]);
  const lookups = useLookups();
  const query = useDashboard(range, project[0] || null);
  const d = query.data;
  return (
    <div className="page">
      <PageHeader
        title="Operations"
        subtitle="How work is flowing through the chapter."
        testId="dashboard-page"
        actions={<Segmented label="Date range" value={range} onChange={setRange} options={RANGES} />}
      />
      <div className="page-toolbar">
        <div className="filter-bar">
          <FilterChip
            label="Project"
            options={projectFilterOptions(lookups.data)}
            value={project}
            onChange={setProject}
            single
          />
          {d && (
            <span className="text-caption text-muted">
              Range {d.range.start.slice(0, 10)} → {d.range.end.slice(0, 10)}
            </span>
          )}
        </div>
      </div>
      <div className="page-body stack" style={{ gap: 20 }}>
        {query.isPending ? (
          <Skeleton lines={8} />
        ) : query.error || !d ? (
          <ErrorState error={query.error} onRetry={() => query.refetch()} />
        ) : (
          <>
            <div className="report-grid report-grid-kpi" data-testid="kpi-grid">
              <KpiCard
                label="Open"
                value={d.totals.open}
                help="Work orders in Draft, Open, In progress or On hold right now."
              />
              <KpiCard
                label="Overdue"
                value={d.totals.overdue}
                tone={d.totals.overdue ? 'bad' : undefined}
                help="Open work orders whose due date has passed."
              />
              <KpiCard label="Blocked" value={d.totals.blocked} help="Open work orders flagged as blocked." />
              <KpiCard label="Due in 7 days" value={d.totals.due_soon} />
              <KpiCard
                label="Created"
                value={d.totals.created}
                help="Work orders created in the selected range."
              />
              <KpiCard
                label="Completed"
                value={d.totals.completed}
                help="Work orders completed in the selected range."
              />
              <KpiCard
                label="On-time completion"
                value={d.on_time_completion_rate == null ? '—' : `${d.on_time_completion_rate}%`}
                help="Completed on or before the due date (no due date counts as on time)."
              />
              <KpiCard
                label="Avg time to complete"
                value={d.average_completion_hours == null ? '—' : `${d.average_completion_hours}h`}
                help="Mean hours from creation to completion."
              />
              <KpiCard label="Hours logged" value={fmtMinutes(Math.round(d.hours_logged * 60))} />
              <KpiCard
                label="Costs"
                value={fmtMoney(d.costs.total)}
                help="Parts, labor, vendor and other costs logged in range."
              />
            </div>
            <div className="grid-2">
              <ReportCard title="Created vs completed" help="Weekly buckets across the selected range.">
                <div className="report-card-body">
                  <ResponsiveContainer width="100%" height={220}>
                    <BarChart
                      data={d.created_vs_completed}
                      margin={{ top: 8, right: 8, left: -18, bottom: 0 }}
                    >
                      <CartesianGrid strokeDasharray="3 3" stroke="#e5ebf1" vertical={false} />
                      <XAxis
                        dataKey="start"
                        tick={{ fontSize: 11 }}
                        tickFormatter={(v: string) => v.slice(5)}
                      />
                      <YAxis tick={{ fontSize: 11 }} allowDecimals={false} />
                      <ChartTooltip cursor={{ fill: '#f3f6f9' }} />
                      <Legend wrapperStyle={{ fontSize: 12 }} />
                      <Bar dataKey="created" name="Created" fill="#0878d1" radius={[4, 4, 0, 0]} />
                      <Bar dataKey="completed" name="Completed" fill="#00a878" radius={[4, 4, 0, 0]} />
                    </BarChart>
                  </ResponsiveContainer>
                </div>
              </ReportCard>
              <ReportCard title="Open work by priority" help="Open work orders grouped by priority.">
                <BarList
                  rows={Object.entries(d.by_priority).map(([k, v]) => ({
                    label: humanize(k),
                    value: v,
                    color: PRIORITY_COLORS[k],
                  }))}
                  emptyText="No open work orders."
                />
              </ReportCard>
            </div>
            <div className="report-grid">
              <ReportCard title="By status">
                <BarList
                  rows={Object.entries(d.by_status).map(([k, v]) => ({
                    label: STATUS_LABEL[k] || k,
                    value: v,
                    color: STATUS_COLORS[k],
                  }))}
                />
              </ReportCard>
              <ReportCard title="Created by work type" help="Work orders created in range, grouped by type.">
                <BarList
                  rows={Object.entries(d.by_work_type).map(([k, v]) => ({ label: humanize(k), value: v }))}
                />
              </ReportCard>
              <ReportCard title="Workload by team" help="Open work orders per team (top 10).">
                <BarList
                  rows={d.workload_by_team.map((w) => ({ label: w.team, value: w.open }))}
                  emptyText="No open work is assigned to a team."
                />
              </ReportCard>
              <ReportCard title="Workload by person" help="Open work orders per assignee (top 10).">
                <BarList
                  rows={d.workload_by_user.map((w) => ({ label: w.name, value: w.open }))}
                  emptyText="No open work is assigned to a person."
                />
              </ReportCard>
              <ReportCard
                title="Repeating vs one-off"
                help="Work orders created in range with and without a recurrence rule."
              >
                <BarList
                  rows={[
                    { label: 'Repeating', value: d.repeating.repeating, color: '#7c5ce7' },
                    { label: 'One-off', value: d.repeating.non_repeating },
                  ]}
                />
              </ReportCard>
              <ReportCard title="Costs by type">
                <BarList
                  rows={[
                    { label: 'Parts', value: d.costs.parts },
                    { label: 'Labor', value: d.costs.labor },
                    { label: 'Vendor', value: d.costs.vendor },
                    { label: 'Other', value: d.costs.other },
                  ].map((r) => ({ ...r, value: Math.round(r.value) }))}
                  emptyText="No costs logged in range."
                />
              </ReportCard>
            </div>
            <p className="text-caption text-muted">
              Need the underlying list? <Link to="/work-orders?filter[due]=overdue">Overdue work orders</Link>{' '}
              · <Link to="/work-orders?filter[blocked]=1">Blocked work orders</Link> · generated{' '}
              {d.generated_at.slice(0, 16).replace('T', ' ')} UTC
            </p>
          </>
        )}
      </div>
    </div>
  );
}
