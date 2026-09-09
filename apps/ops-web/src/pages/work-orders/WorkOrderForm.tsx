import { useEffect, useState } from 'react';
import { Controller, useForm } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import { api, ApiError } from '@/api/client';
import { keys, useInvalidate, useLookups } from '@/api/hooks';
import { useCurrentSession } from '@/auth/SessionProvider';
import {
  workOrderCreateSchema,
  WO_PRIORITIES,
  WO_WORK_TYPES,
  type WorkOrderCreateInput,
} from '@/contracts/schemas';
import type { WorkOrderDetail } from '@/contracts/types';
import { humanize, toLocalInput } from '@/lib/format';
import {
  assetOptions,
  categoryOptions,
  locationOptions,
  projectOptions,
  teamOptions,
  userOptions,
} from '@/lib/options';
import {
  Button,
  Checkbox,
  Picker,
  Segmented,
  SelectField,
  SideSheet,
  TextArea,
  TextField,
  useToast,
} from '@/ui';

type Mode = 'create' | 'edit';

/** Fields that PATCH /work-orders/:id accepts (assignees and watchers have their own endpoints). */
const EDITABLE = [
  'title',
  'description',
  'project_id',
  'location_id',
  'primary_asset_id',
  'asset_ids',
  'team_id',
  'estimated_minutes',
  'due_at',
  'start_at',
  'recurrence',
  'work_type',
  'priority',
  'category_ids',
  'budget_code',
  'parent_completion_policy',
] as const;

function initialValues(wo?: WorkOrderDetail, defaults?: Partial<WorkOrderCreateInput>): WorkOrderCreateInput {
  return {
    title: wo?.title || '',
    description: wo?.description || '',
    project_id: wo?.project?.id || defaults?.project_id || '',
    location_id: wo?.location?.id || '',
    primary_asset_id: wo?.primary_asset?.id || '',
    asset_ids: wo?.related_assets.map((a) => a.id) || [],
    assignee_user_ids: [],
    assignee_team_ids: [],
    team_id: wo?.team?.id || '',
    estimated_minutes: wo?.estimated_minutes ?? undefined,
    due_at: toLocalInput(wo?.due_at),
    start_at: toLocalInput(wo?.start_at),
    recurrence: wo?.recurrence || null,
    work_type: wo?.work_type || 'PROJECT',
    priority: wo?.priority || 'NONE',
    category_ids: wo?.categories.map((c) => c.id) || [],
    budget_code: wo?.budget_code || '',
    watcher_user_ids: [],
    parent_work_order_id: '',
    status: 'OPEN',
    parent_completion_policy: wo?.parent_completion_policy || 'manual',
    ...(defaults || {}),
  };
}

export function WorkOrderForm({
  mode,
  workOrder,
  defaults,
  onClose,
  onCreated,
  onSaved,
}: {
  mode: Mode;
  workOrder?: WorkOrderDetail;
  defaults?: Partial<WorkOrderCreateInput>;
  onClose: () => void;
  onCreated?: (wo: WorkOrderDetail) => void;
  onSaved?: (wo: WorkOrderDetail) => void;
}) {
  const toast = useToast();
  const lookups = useLookups();
  const invalidate = useInvalidate();
  const { can } = useCurrentSession();
  const [banner, setBanner] = useState<string | null>(null);
  const [createAnother, setCreateAnother] = useState(false);
  const [recurring, setRecurring] = useState(!!workOrder?.recurrence);

  const form = useForm<WorkOrderCreateInput>({
    resolver: zodResolver(workOrderCreateSchema),
    defaultValues: initialValues(workOrder, defaults),
  });
  const projectId = form.watch('project_id');
  const primaryAsset = form.watch('primary_asset_id');
  const priority = form.watch('priority');

  useEffect(() => {
    if (!recurring) form.setValue('recurrence', null);
    else if (!form.getValues('recurrence'))
      form.setValue('recurrence', { frequency: 'weekly', interval: 1, mode: 'fixed' });
  }, [recurring, form]);

  const submit = async (values: WorkOrderCreateInput, statusOverride?: 'DRAFT' | 'OPEN') => {
    setBanner(null);
    const data = workOrderCreateSchema.parse({ ...values, status: statusOverride || values.status });
    try {
      if (mode === 'create') {
        const created = await api<WorkOrderDetail>('/work-orders', { method: 'POST', body: data });
        invalidate(['work-orders'], keys.lookups, keys.setup, ['projects'], ['project'], ['project-health']);
        toast.success(`Work order #${created.number} created`, created.title);
        if (createAnother) {
          form.reset(
            initialValues(undefined, {
              ...defaults,
              project_id: data.project_id || defaults?.project_id,
              team_id: data.team_id,
              location_id: data.location_id,
            }),
          );
          return;
        }
        onCreated?.(created);
      } else if (workOrder) {
        const body: Record<string, unknown> = {};
        const dirty = form.formState.dirtyFields as Record<string, unknown>;
        for (const key of EDITABLE) {
          if (dirty[key]) body[key] = (data as Record<string, unknown>)[key] ?? null;
        }
        if (Object.keys(body).length === 0) {
          onClose();
          return;
        }
        const saved = await api<WorkOrderDetail>(`/work-orders/${workOrder.id}`, { method: 'PATCH', body });
        invalidate(['work-orders'], keys.workOrder(workOrder.id), ['project-health']);
        toast.success('Work order updated');
        onSaved?.(saved);
      }
    } catch (err) {
      if (err instanceof ApiError) {
        const fields = err.fieldErrors();
        Object.entries(fields).forEach(([field, message]) =>
          form.setError(field as keyof WorkOrderCreateInput, { message }),
        );
        if (!Object.keys(fields).length) setBanner(err.message);
      } else setBanner('Could not save. Check your connection and try again.');
    }
  };

  const errors = form.formState.errors;
  const busy = form.formState.isSubmitting;
  const canSetCritical =
    mode === 'create' ? can('work_order.set_critical') : (workOrder?.permissions.set_critical ?? false);

  return (
    <SideSheet
      open
      onClose={onClose}
      title={mode === 'create' ? 'New work order' : `Edit #${workOrder?.number}`}
      subtitle={
        mode === 'create'
          ? 'Title is the only required field. Everything else can be added later.'
          : undefined
      }
      testId="work-order-form"
      footer={
        <>
          {mode === 'create' && (
            <Checkbox
              label="Create another"
              checked={createAnother}
              onChange={(e) => setCreateAnother(e.target.checked)}
            />
          )}
          <span className="grow" />
          <Button variant="ghost" onClick={onClose} disabled={busy}>
            Cancel
          </Button>
          {mode === 'create' && (
            <Button
              variant="secondary"
              onClick={form.handleSubmit((v) => submit(v, 'DRAFT'))}
              disabled={busy}
              data-testid="save-draft"
            >
              Save as draft
            </Button>
          )}
          <Button
            variant="primary"
            onClick={form.handleSubmit((v) => submit(v, mode === 'create' ? 'OPEN' : undefined))}
            loading={busy}
            data-testid="work-order-submit"
          >
            {mode === 'create' ? 'Create work order' : 'Save changes'}
          </Button>
        </>
      }
    >
      <form
        className="stack"
        style={{ gap: 20 }}
        onSubmit={form.handleSubmit((v) => submit(v, mode === 'create' ? 'OPEN' : undefined))}
        noValidate
      >
        {banner && (
          <div className="form-banner form-banner-error" role="alert">
            {banner}
          </div>
        )}
        <TextField
          label="Title"
          required
          placeholder="What needs to be done?"
          error={errors.title?.message}
          {...form.register('title')}
          data-autofocus
          data-testid="wo-title"
          maxLength={220}
        />
        <TextArea
          label="Description"
          placeholder="Steps, acceptance criteria, links to CAD or issues…"
          error={errors.description?.message}
          {...form.register('description')}
          rows={4}
        />

        <div className="form-section">
          <div className="form-section-title">Priority &amp; type</div>
          <div className="field">
            <span className="field-label">Priority</span>
            <Segmented
              label="Priority"
              value={priority || 'NONE'}
              onChange={(v) => form.setValue('priority', v, { shouldDirty: true })}
              options={WO_PRIORITIES.filter(
                (p) => p !== 'CRITICAL' || canSetCritical || priority === 'CRITICAL',
              ).map((p) => ({ value: p, label: humanize(p) }))}
            />
            {!canSetCritical && (
              <div className="field-hint">
                Critical priority can be set by safety officers, leads and admins.
              </div>
            )}
          </div>
          <SelectField
            label="Work type"
            options={WO_WORK_TYPES.map((t) => ({ value: t, label: humanize(t) }))}
            error={errors.work_type?.message}
            {...form.register('work_type')}
          />
          <Controller
            control={form.control}
            name="category_ids"
            render={({ field }) => (
              <Picker
                label="Categories"
                options={categoryOptions(lookups.data)}
                value={field.value || []}
                onChange={field.onChange}
                placeholder="Add categories"
              />
            )}
          />
        </div>

        <div className="form-section">
          <div className="form-section-title">Where it belongs</div>
          <Controller
            control={form.control}
            name="project_id"
            render={({ field, fieldState }) => (
              <Picker
                label="Project"
                multiple={false}
                options={projectOptions(lookups.data)}
                value={field.value ? [field.value] : []}
                onChange={(v) => field.onChange(v[0] || '')}
                placeholder="No project"
                error={fieldState.error?.message}
                id="wo-project"
              />
            )}
          />
          <Controller
            control={form.control}
            name="team_id"
            render={({ field, fieldState }) => (
              <Picker
                label="Team"
                multiple={false}
                options={teamOptions(lookups.data, projectId || null)}
                value={field.value ? [field.value] : []}
                onChange={(v) => field.onChange(v[0] || '')}
                placeholder="No team"
                error={fieldState.error?.message}
                id="wo-team"
              />
            )}
          />
          <div className="field-row">
            <Controller
              control={form.control}
              name="location_id"
              render={({ field, fieldState }) => (
                <Picker
                  label="Location"
                  multiple={false}
                  options={locationOptions(lookups.data)}
                  value={field.value ? [field.value] : []}
                  onChange={(v) => field.onChange(v[0] || '')}
                  placeholder="Default location"
                  error={fieldState.error?.message}
                />
              )}
            />
            <Controller
              control={form.control}
              name="primary_asset_id"
              render={({ field, fieldState }) => (
                <Picker
                  label="Asset"
                  multiple={false}
                  options={assetOptions(lookups.data, projectId || null)}
                  value={field.value ? [field.value] : []}
                  onChange={(v) => field.onChange(v[0] || '')}
                  placeholder="No asset"
                  error={fieldState.error?.message}
                  id="wo-asset"
                />
              )}
            />
          </div>
          {primaryAsset && (
            <Controller
              control={form.control}
              name="asset_ids"
              render={({ field }) => (
                <Picker
                  label="Related assets"
                  options={assetOptions(lookups.data, projectId || null).filter(
                    (o) => o.value !== primaryAsset,
                  )}
                  value={field.value || []}
                  onChange={field.onChange}
                  placeholder="Other assets involved"
                />
              )}
            />
          )}
        </div>

        {mode === 'create' && (
          <div className="form-section">
            <div className="form-section-title">Assign</div>
            <Controller
              control={form.control}
              name="assignee_user_ids"
              render={({ field, fieldState }) => (
                <Picker
                  label="People"
                  options={userOptions(lookups.data)}
                  value={field.value || []}
                  onChange={field.onChange}
                  placeholder="Assign people"
                  error={fieldState.error?.message}
                  id="wo-assignees"
                />
              )}
            />
            <Controller
              control={form.control}
              name="assignee_team_ids"
              render={({ field }) => (
                <Picker
                  label="Teams"
                  options={teamOptions(lookups.data, projectId || null)}
                  value={field.value || []}
                  onChange={field.onChange}
                  placeholder="Assign a whole team"
                />
              )}
            />
            <Controller
              control={form.control}
              name="watcher_user_ids"
              render={({ field }) => (
                <Picker
                  label="Watchers"
                  options={userOptions(lookups.data)}
                  value={field.value || []}
                  onChange={field.onChange}
                  placeholder="People to keep in the loop"
                />
              )}
            />
          </div>
        )}

        <div className="form-section">
          <div className="form-section-title">Schedule</div>
          <div className="field-row">
            <TextField
              label="Start"
              type="datetime-local"
              error={errors.start_at?.message}
              {...form.register('start_at')}
            />
            <TextField
              label="Due"
              type="datetime-local"
              error={errors.due_at?.message}
              {...form.register('due_at')}
              data-testid="wo-due"
            />
          </div>
          <div className="field-row">
            <TextField
              label="Estimated time (minutes)"
              type="number"
              min={0}
              step={5}
              placeholder="e.g. 90"
              error={errors.estimated_minutes?.message}
              {...form.register('estimated_minutes', { valueAsNumber: true })}
            />
            <TextField
              label="Budget code"
              placeholder="e.g. ASME-CCR-27"
              error={errors.budget_code?.message}
              {...form.register('budget_code')}
            />
          </div>
          <Checkbox
            label="Repeats on a schedule"
            checked={recurring}
            onChange={(e) => setRecurring(e.target.checked)}
          />
          {recurring && (
            <div className="field-row">
              <Controller
                control={form.control}
                name="recurrence"
                render={({ field }) => (
                  <>
                    <SelectField
                      label="Every"
                      options={[
                        { value: 'daily', label: 'Day(s)' },
                        { value: 'weekly', label: 'Week(s)' },
                        { value: 'monthly', label: 'Month(s)' },
                      ]}
                      value={field.value?.frequency || 'weekly'}
                      onChange={(e) =>
                        field.onChange({
                          ...(field.value || { interval: 1, mode: 'fixed' }),
                          frequency: e.target.value,
                        })
                      }
                    />
                    <TextField
                      label="Interval"
                      type="number"
                      min={1}
                      max={52}
                      value={field.value?.interval ?? 1}
                      onChange={(e) =>
                        field.onChange({
                          ...(field.value || { frequency: 'weekly', mode: 'fixed' }),
                          interval: Number(e.target.value) || 1,
                        })
                      }
                    />
                    <SelectField
                      label="Next occurrence"
                      options={[
                        { value: 'fixed', label: 'Fixed: from the due date' },
                        { value: 'floating', label: 'Floating: from completion' },
                      ]}
                      value={field.value?.mode || 'fixed'}
                      onChange={(e) =>
                        field.onChange({
                          ...(field.value || { frequency: 'weekly', interval: 1 }),
                          mode: e.target.value,
                        })
                      }
                      wrapClassName="grid-span"
                    />
                  </>
                )}
              />
            </div>
          )}
        </div>
      </form>
    </SideSheet>
  );
}
