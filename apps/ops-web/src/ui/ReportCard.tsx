import type { ReactNode } from 'react';
import { HelpCircle } from 'lucide-react';
import { Tooltip } from './Popover';

export function ReportCard({
  title,
  help,
  children,
  actions,
}: {
  title: ReactNode;
  help?: string;
  children: ReactNode;
  actions?: ReactNode;
}) {
  return (
    <section className="report-card" aria-label={typeof title === 'string' ? title : undefined}>
      <div className="report-card-title">
        <span className="row" style={{ gap: 6 }}>
          {title}
          {help && (
            <Tooltip text={help}>
              <button
                type="button"
                className="help link-button"
                aria-label={`About ${typeof title === 'string' ? title : 'this metric'}: ${help}`}
              >
                <HelpCircle size={14} />
              </button>
            </Tooltip>
          )}
        </span>
        {actions}
      </div>
      {children}
    </section>
  );
}

export function KpiCard({
  label,
  value,
  help,
  delta,
  tone,
}: {
  label: string;
  value: ReactNode;
  help?: string;
  delta?: string;
  tone?: 'good' | 'bad';
}) {
  return (
    <ReportCard title={label} help={help}>
      <div className="kpi-value" data-testid={`kpi-${label.toLowerCase().replace(/[^a-z0-9]+/g, '-')}`}>
        {value}
      </div>
      {delta && <div className={`kpi-delta ${tone || ''}`}>{delta}</div>}
    </ReportCard>
  );
}

export function BarList({
  rows,
  color = 'var(--color-primary)',
  emptyText = 'No data in this range.',
}: {
  rows: { label: string; value: number; color?: string }[];
  color?: string;
  emptyText?: string;
}) {
  const max = Math.max(1, ...rows.map((r) => r.value));
  if (rows.length === 0 || rows.every((r) => r.value === 0))
    return <div className="text-muted text-label">{emptyText}</div>;
  return (
    <div className="bar-list">
      {rows.map((r) => (
        <div key={r.label} className="bar-list-row">
          <div>
            <div className="row-between">
              <span className="truncate">{r.label}</span>
            </div>
            <div className="bar-list-track">
              <div
                className="bar-list-fill"
                style={{ width: `${(r.value / max) * 100}%`, background: r.color || color }}
              />
            </div>
          </div>
          <div
            className="num"
            style={{ textAlign: 'right', fontVariantNumeric: 'tabular-nums', fontWeight: 600 }}
          >
            {r.value}
          </div>
        </div>
      ))}
    </div>
  );
}
