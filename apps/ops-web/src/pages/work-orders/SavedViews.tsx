import { useState, type ReactNode } from 'react';
import { useForm, Controller } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import { ChevronDown, Star, Trash2 } from 'lucide-react';
import { api, ApiError } from '@/api/client';
import { keys, useApiMutation, useLookups, useSavedFilters } from '@/api/hooks';
import { savedFilterSchema, type SavedFilterInput } from '@/contracts/schemas';
import type { SavedFilter } from '@/contracts/types';
import {
  Button,
  Checkbox,
  ConfirmDialog,
  Dialog,
  Picker,
  Popover,
  SelectField,
  TextField,
  useToast,
} from '@/ui';

/**
 * Saved view selector (spec §7.4): apply a saved filter set, save the current one,
 * share with a team or the chapter, or delete your own.
 */
export function SavedViews({
  entityType,
  filters,
  sort,
  viewType,
  onApply,
  canShare,
  icon,
}: {
  entityType: 'work_order' | 'project' | 'asset';
  filters: Record<string, string[]>;
  sort: string;
  viewType: 'panel' | 'table';
  onApply: (filters: Record<string, string[]>, sort: string | null, viewType: string) => void;
  canShare: boolean;
  icon?: ReactNode;
}) {
  const toast = useToast();
  const views = useSavedFilters(entityType);
  const lookups = useLookups();
  const [saving, setSaving] = useState(false);
  const [deleting, setDeleting] = useState<SavedFilter | null>(null);
  const create = useApiMutation(
    (body: Record<string, unknown>) => api<SavedFilter>('/saved-filters', { method: 'POST', body }),
    [keys.savedFilters(entityType)],
  );
  const remove = useApiMutation(
    (id: string) => api(`/saved-filters/${id}`, { method: 'DELETE' }),
    [keys.savedFilters(entityType)],
  );
  const hasFilters = Object.keys(filters).length > 0 || !!sort;

  const form = useForm<SavedFilterInput>({
    resolver: zodResolver(savedFilterSchema),
    defaultValues: { name: '', visibility: 'private', team_id: '', is_default: false },
  });
  const visibility = form.watch('visibility');

  const all = [...(views.data?.mine || []), ...(views.data?.shared || [])];

  return (
    <>
      <Popover
        align="right"
        label="Saved views"
        trigger={
          <button type="button" className="filter-chip" data-testid="saved-views">
            {icon} Views {all.length > 0 && `(${all.length})`} <ChevronDown />
          </button>
        }
      >
        {(close) => (
          <div className="menu" style={{ width: 300 }}>
            {views.data?.mine.length ? <div className="menu-heading">My views</div> : null}
            {views.data?.mine.map((v) => (
              <ViewRow
                key={v.id}
                view={v}
                onApply={() => {
                  onApply(v.filters, v.sort, v.view_type);
                  close();
                }}
                onDelete={() => setDeleting(v)}
              />
            ))}
            {views.data?.shared.length ? <div className="menu-heading">Shared with me</div> : null}
            {views.data?.shared.map((v) => (
              <ViewRow
                key={v.id}
                view={v}
                onApply={() => {
                  onApply(v.filters, v.sort, v.view_type);
                  close();
                }}
              />
            ))}
            {all.length === 0 && (
              <div className="picker-empty">No saved views yet. Set some filters, then save them here.</div>
            )}
            <div className="menu-separator" />
            <button
              type="button"
              className="menu-item"
              disabled={!hasFilters}
              onClick={() => {
                close();
                form.reset({ name: '', visibility: 'private', team_id: '', is_default: false });
                setSaving(true);
              }}
              data-testid="save-view"
            >
              <Star /> Save current filters as a view
            </button>
            {!hasFilters && (
              <div className="text-caption text-muted" style={{ padding: '0 10px 8px' }}>
                Choose at least one filter or sort first.
              </div>
            )}
          </div>
        )}
      </Popover>

      <Dialog
        open={saving}
        onClose={() => setSaving(false)}
        title="Save view"
        description="Saved views remember filters, sort and layout. Shared views are visible to the team or the whole chapter."
        footer={
          <>
            <Button variant="ghost" onClick={() => setSaving(false)}>
              Cancel
            </Button>
            <Button
              variant="primary"
              loading={create.isPending}
              onClick={form.handleSubmit(async (values) => {
                const data = savedFilterSchema.parse(values);
                try {
                  await create.mutateAsync({
                    entity_type: entityType,
                    name: data.name,
                    visibility: data.visibility,
                    team_id: data.team_id,
                    filters,
                    sort: sort || null,
                    view_type: viewType,
                    is_default: data.is_default,
                  });
                  toast.success('View saved');
                  setSaving(false);
                } catch (err) {
                  if (err instanceof ApiError) {
                    const fields = err.fieldErrors();
                    Object.entries(fields).forEach(([f, m]) =>
                      form.setError(f as keyof SavedFilterInput, { message: m }),
                    );
                    if (!Object.keys(fields).length) toast.error('Could not save view', err.message);
                  }
                }
              })}
              data-testid="save-view-submit"
            >
              Save view
            </Button>
          </>
        }
      >
        <div className="stack" style={{ gap: 16 }}>
          <TextField
            label="Name"
            required
            error={form.formState.errors.name?.message}
            {...form.register('name')}
            data-autofocus
            placeholder="e.g. My overdue work"
          />
          <SelectField
            label="Visibility"
            options={[
              { value: 'private', label: 'Only me' },
              { value: 'team', label: 'A team', disabled: !canShare },
              { value: 'chapter', label: 'Whole chapter', disabled: !canShare },
            ]}
            hint={!canShare ? 'Sharing views needs the "share saved filters" permission.' : undefined}
            {...form.register('visibility')}
          />
          {visibility === 'team' && (
            <Controller
              control={form.control}
              name="team_id"
              render={({ field, fieldState }) => (
                <Picker
                  label="Team"
                  multiple={false}
                  options={(lookups.data?.teams || []).map((t) => ({
                    value: t.id,
                    label: t.name,
                    color: t.color,
                  }))}
                  value={field.value ? [field.value] : []}
                  onChange={(v) => field.onChange(v[0] || '')}
                  error={fieldState.error?.message}
                  required
                />
              )}
            />
          )}
          <Checkbox label="Use as my default view" {...form.register('is_default')} />
        </div>
      </Dialog>

      <ConfirmDialog
        open={!!deleting}
        onClose={() => setDeleting(null)}
        title="Delete view?"
        body={`"${deleting?.name}" will be removed${deleting?.visibility !== 'private' ? ' for everyone it is shared with' : ''}.`}
        confirmLabel="Delete"
        danger
        loading={remove.isPending}
        onConfirm={async () => {
          if (!deleting) return;
          await remove.mutateAsync(deleting.id);
          toast.success('View deleted');
          setDeleting(null);
        }}
      />
    </>
  );
}

function ViewRow({
  view,
  onApply,
  onDelete,
}: {
  view: SavedFilter;
  onApply: () => void;
  onDelete?: () => void;
}) {
  return (
    <div className="row" style={{ gap: 0 }}>
      <button
        type="button"
        className="menu-item"
        style={{ flex: 1 }}
        onClick={onApply}
        data-testid="saved-view-item"
      >
        {view.is_default && <Star size={14} />}
        <span style={{ flex: 1 }}>
          {view.name}
          <span className="text-caption text-muted" style={{ display: 'block' }}>
            {view.visibility === 'private' ? 'Only me' : view.visibility === 'team' ? 'Team' : 'Chapter'}
            {!view.is_mine && view.owner ? ` · ${view.owner.name}` : ''}
          </span>
        </span>
      </button>
      {onDelete && view.is_mine && (
        <button
          type="button"
          className="btn btn-ghost btn-icon btn-sm"
          aria-label={`Delete view ${view.name}`}
          onClick={onDelete}
        >
          <Trash2 size={14} />
        </button>
      )}
    </div>
  );
}
