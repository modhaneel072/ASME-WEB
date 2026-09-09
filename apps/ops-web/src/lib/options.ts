import type { Lookups } from '@/contracts/types';
import type { PickerOption } from '@/ui/Picker';
import type { FilterOption } from '@/ui/FilterChip';
import { humanize, STATUS_LABEL } from './format';

export const WO_STATUS_OPEN = ['DRAFT', 'OPEN', 'IN_PROGRESS', 'ON_HOLD'] as const;
export const WO_STATUS_CLOSED = ['DONE', 'CANCELED', 'SKIPPED'] as const;
export const PRIORITIES = ['NONE', 'LOW', 'MEDIUM', 'HIGH', 'CRITICAL'] as const;
export const WORK_TYPES = [
  'REACTIVE',
  'PREVENTIVE',
  'PROJECT',
  'EVENT',
  'INSPECTION',
  'SAFETY',
  'PROCUREMENT',
  'DOCUMENTATION',
] as const;
export const ASSET_STATUSES = [
  'ONLINE',
  'OFFLINE_PLANNED',
  'OFFLINE_UNPLANNED',
  'DO_NOT_TRACK',
  'RETIRED',
] as const;

export const statusOptions = (keys: readonly string[]): FilterOption[] =>
  keys.map((k) => ({ value: k, label: STATUS_LABEL[k] || humanize(k) }));
export const priorityOptions: FilterOption[] = PRIORITIES.map((p) => ({ value: p, label: humanize(p) }));
export const workTypeOptions: FilterOption[] = WORK_TYPES.map((t) => ({ value: t, label: humanize(t) }));
export const dueOptions: FilterOption[] = [
  { value: 'overdue', label: 'Overdue' },
  { value: 'today', label: 'Due today' },
  { value: 'week', label: 'Next 7 days' },
  { value: 'month', label: 'Next 30 days' },
  { value: 'none', label: 'No due date' },
  { value: 'has', label: 'Has due date' },
];

export function userOptions(lookups: Lookups | undefined): PickerOption<number>[] {
  return (lookups?.users || []).map((u) => ({ value: u.id, label: u.name, hint: u.role_name || u.email }));
}

export function userFilterOptions(lookups: Lookups | undefined, withSpecial = true): FilterOption[] {
  const base = (lookups?.users || []).map((u) => ({ value: String(u.id), label: u.name }));
  return withSpecial
    ? [{ value: 'me', label: 'Assigned to me' }, { value: 'unassigned', label: 'Unassigned' }, ...base]
    : base;
}

export function teamOptions(lookups: Lookups | undefined, projectId?: string | null): PickerOption<string>[] {
  const teams = lookups?.teams || [];
  return teams
    .filter((t) => !projectId || !t.project_id || t.project_id === projectId)
    .map((t) => ({
      value: t.id,
      label: t.name,
      color: t.color,
      group: t.project_id
        ? lookups?.projects.find((p) => p.id === t.project_id)?.name || 'Project teams'
        : 'Chapter teams',
    }));
}

export const teamFilterOptions = (lookups: Lookups | undefined): FilterOption[] =>
  (lookups?.teams || []).map((t) => ({ value: t.id, label: t.name, color: t.color }));

export function locationOptions(lookups: Lookups | undefined): PickerOption<string>[] {
  const locations = lookups?.locations || [];
  const byId = new Map(locations.map((l) => [l.id, l]));
  const path = (id: string): string => {
    const loc = byId.get(id);
    if (!loc) return '';
    return loc.parent_location_id ? `${path(loc.parent_location_id)} › ${loc.name}` : loc.name;
  };
  return locations.map((l) => ({
    value: l.id,
    label: l.name,
    hint: l.parent_location_id ? path(l.parent_location_id) : l.is_default ? 'Default' : undefined,
  }));
}

export const locationFilterOptions = (lookups: Lookups | undefined): FilterOption[] =>
  (lookups?.locations || []).map((l) => ({ value: l.id, label: l.name }));

export function categoryOptions(lookups: Lookups | undefined): PickerOption<string>[] {
  return (lookups?.categories || []).map((c) => ({ value: c.id, label: c.name, color: c.color }));
}

export const categoryFilterOptions = (lookups: Lookups | undefined): FilterOption[] =>
  (lookups?.categories || []).map((c) => ({ value: c.id, label: c.name, color: c.color }));

export function projectOptions(lookups: Lookups | undefined): PickerOption<string>[] {
  return (lookups?.projects || []).map((p) => ({ value: p.id, label: p.name, hint: p.code }));
}

export const projectFilterOptions = (lookups: Lookups | undefined, withNone = false): FilterOption[] => [
  ...(withNone ? [{ value: 'none', label: 'No project' }] : []),
  ...(lookups?.projects || []).map((p) => ({ value: p.id, label: `${p.name} (${p.code})` })),
];

export function assetOptions(
  lookups: Lookups | undefined,
  projectId?: string | null,
): PickerOption<string>[] {
  const assets = lookups?.assets || [];
  const byId = new Map(assets.map((a) => [a.id, a]));
  return assets
    .filter((a) => !projectId || !a.project_id || a.project_id === projectId)
    .map((a) => ({
      value: a.id,
      label: a.name,
      hint:
        [a.code, a.parent_asset_id ? byId.get(a.parent_asset_id)?.name : null].filter(Boolean).join(' · ') ||
        undefined,
      group: a.project_id
        ? lookups?.projects.find((p) => p.id === a.project_id)?.name || 'Project assets'
        : 'Shared equipment',
    }));
}

export const assetFilterOptions = (lookups: Lookups | undefined): FilterOption[] =>
  (lookups?.assets || []).map((a) => ({ value: a.id, label: a.code ? `${a.name} (${a.code})` : a.name }));
