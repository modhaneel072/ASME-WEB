import { format, formatDistanceToNowStrict, isToday, isTomorrow, isYesterday, parseISO } from 'date-fns';

export function parse(value: string | null | undefined): Date | null {
  if (!value) return null;
  const d =
    value.length <= 10
      ? parseISO(value)
      : new Date(value.endsWith('Z') || /[+-]\d\d:\d\d$/.test(value) ? value : `${value}Z`);
  return Number.isNaN(d.getTime()) ? null : d;
}

export function fmtDate(value: string | null | undefined, withTime = false): string {
  const d = parse(value);
  if (!d) return '—';
  if (withTime) return format(d, 'MMM d, yyyy h:mm a');
  return format(d, 'MMM d, yyyy');
}

export function fmtDue(value: string | null | undefined): string {
  const d = parse(value);
  if (!d) return 'No due date';
  if (isToday(d)) return `Today ${format(d, 'h:mm a')}`;
  if (isTomorrow(d)) return `Tomorrow ${format(d, 'h:mm a')}`;
  if (isYesterday(d)) return `Yesterday ${format(d, 'h:mm a')}`;
  return format(d, 'MMM d, h:mm a');
}

export function fmtRelative(value: string | null | undefined): string {
  const d = parse(value);
  if (!d) return '';
  return `${formatDistanceToNowStrict(d)} ago`;
}

export function fmtMinutes(minutes: number | null | undefined): string {
  if (!minutes) return '0m';
  const h = Math.floor(minutes / 60);
  const m = minutes % 60;
  if (h && m) return `${h}h ${m}m`;
  if (h) return `${h}h`;
  return `${m}m`;
}

export function fmtMoney(value: number | null | undefined): string {
  if (value === null || value === undefined) return '—';
  return new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: 'USD',
    maximumFractionDigits: 2,
  }).format(value);
}

export function fmtBytes(bytes: number): string {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

export function toLocalInput(value: string | null | undefined): string {
  const d = parse(value);
  return d ? format(d, "yyyy-MM-dd'T'HH:mm") : '';
}

export function humanize(value: string | null | undefined): string {
  if (!value) return '';
  return value
    .replace(/_/g, ' ')
    .toLowerCase()
    .replace(/(^|\s)\S/g, (c) => c.toUpperCase());
}

export const STATUS_LABEL: Record<string, string> = {
  DRAFT: 'Draft',
  OPEN: 'Open',
  IN_PROGRESS: 'In progress',
  ON_HOLD: 'On hold',
  DONE: 'Done',
  CANCELED: 'Canceled',
  SKIPPED: 'Skipped',
  ONLINE: 'Online',
  OFFLINE_PLANNED: 'Offline (planned)',
  OFFLINE_UNPLANNED: 'Offline (unplanned)',
  DO_NOT_TRACK: 'Not tracked',
  RETIRED: 'Retired',
};
