import type { AuditEvent, StatusHistoryEntry } from '@/contracts/types';
import { STATUS_LABEL, fmtDate, humanize } from '@/lib/format';
import { cn } from '@/lib/cn';

export interface TimelineItem {
  id: string;
  title: string;
  meta: string;
  note?: string | null;
  tone?: 'default' | 'muted' | 'success' | 'danger';
}

export function ActivityTimeline({
  items,
  emptyText = 'No activity yet.',
}: {
  items: TimelineItem[];
  emptyText?: string;
}) {
  if (items.length === 0) return <div className="text-muted text-label">{emptyText}</div>;
  return (
    <ol className="timeline" style={{ listStyle: 'none', margin: 0 }}>
      {items.map((item) => (
        <li key={item.id} className="timeline-item">
          <span
            className={cn('timeline-dot', item.tone && item.tone !== 'default' && item.tone)}
            aria-hidden="true"
          />
          <div className="timeline-title">{item.title}</div>
          <div className="timeline-meta">{item.meta}</div>
          {item.note && <div className="timeline-note">{item.note}</div>}
        </li>
      ))}
    </ol>
  );
}

const TONE_FOR_STATUS: Record<string, TimelineItem['tone']> = {
  DONE: 'success',
  CANCELED: 'danger',
  SKIPPED: 'danger',
  ON_HOLD: 'muted',
};

export function statusHistoryItems(entries: StatusHistoryEntry[]): TimelineItem[] {
  return [...entries]
    .sort((a, b) => (a.changed_at < b.changed_at ? 1 : -1))
    .map((h) => ({
      id: h.id,
      title: h.from_status
        ? `${STATUS_LABEL[h.from_status] || h.from_status} → ${STATUS_LABEL[h.to_status] || h.to_status}`
        : `Created as ${STATUS_LABEL[h.to_status] || h.to_status}`,
      meta: `${h.changed_by?.name || 'System'} · ${fmtDate(h.changed_at, true)}`,
      note: h.note,
      tone: TONE_FOR_STATUS[h.to_status] || 'default',
    }));
}

/** Human sentence for an audit event; unknown types fall back to the raw key. */
export function describeEvent(event: AuditEvent): string {
  const meta = event.metadata || {};
  const [entity, action] = event.event_type.split('.');
  const label = humanize(entity);
  switch (action) {
    case 'created':
      return `${label} created${meta.title ? `: ${meta.title}` : ''}`;
    case 'updated': {
      const changed = meta.changed && typeof meta.changed === 'object' ? Object.keys(meta.changed) : [];
      return changed.length
        ? `${label} updated (${changed.map((c) => humanize(c).toLowerCase()).join(', ')})`
        : `${label} updated`;
    }
    case 'status_changed':
      return `${label} status ${meta.from ? `${STATUS_LABEL[meta.from] || meta.from} → ` : ''}${STATUS_LABEL[meta.to] || meta.to || ''}`.trim();
    case 'assigned':
    case 'assignees_changed':
      return `${label} assignees changed`;
    case 'archived':
      return `${label} archived`;
    case 'deleted':
      return `${label} deleted`;
    case 'commented':
      return `Comment added`;
    case 'file_uploaded':
      return `File uploaded${meta.name ? `: ${meta.name}` : ''}`;
    case 'time_logged':
      return `Time logged${meta.minutes ? `: ${meta.minutes} min` : ''}`;
    case 'cost_logged':
      return `Cost logged${meta.amount ? `: $${meta.amount}` : ''}`;
    default:
      return `${label} ${humanize(action || '').toLowerCase()}`.trim();
  }
}

export function auditItems(events: AuditEvent[]): TimelineItem[] {
  return events.map((e) => ({
    id: e.id,
    title: describeEvent(e),
    meta: `${e.actor?.name || 'System'} · ${fmtDate(e.occurred_at, true)}`,
    note: typeof e.metadata?.note === 'string' ? e.metadata.note : null,
    tone:
      e.event_type.endsWith('archived') || e.event_type.endsWith('deleted')
        ? 'danger'
        : e.event_type.endsWith('created')
          ? 'success'
          : 'default',
  }));
}
