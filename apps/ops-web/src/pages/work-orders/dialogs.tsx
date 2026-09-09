import { useState } from 'react';
import { Controller, useFieldArray, useForm } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import { Plus, Trash2 } from 'lucide-react';
import { api, ApiError } from '@/api/client';
import { useLookups } from '@/api/hooks';
import { completeSchema, subWorkOrderSchema, WO_PRIORITIES, type CompleteInput } from '@/contracts/schemas';
import type { WorkOrderDetail } from '@/contracts/types';
import { humanize, STATUS_LABEL } from '@/lib/format';
import { ASSET_STATUSES, teamOptions, userOptions } from '@/lib/options';
import { Button, Dialog, Picker, SelectField, TextArea, TextField, useToast } from '@/ui';
import { z } from 'zod';

type Done = (wo: WorkOrderDetail) => void;

function useFieldErrors<T extends Record<string, unknown>>(
  setError: (name: keyof T & string, error: { message: string }) => void,
  setBanner: (m: string | null) => void,
) {
  return (err: unknown) => {
    if (err instanceof ApiError) {
      const fields = err.fieldErrors();
      Object.entries(fields).forEach(([f, m]) => setError(f as keyof T & string, { message: m }));
      if (!Object.keys(fields).length) setBanner(err.message);
    } else setBanner('Something went wrong. Try again.');
  };
}

/** Complete with note, time, costs, asset status and optional follow-up (spec §8.4). */
export function CompleteDialog({
  wo,
  open,
  onClose,
  onDone,
}: {
  wo: WorkOrderDetail;
  open: boolean;
  onClose: () => void;
  onDone: Done;
}) {
  const toast = useToast();
  const [banner, setBanner] = useState<string | null>(null);
  const form = useForm<CompleteInput>({
    resolver: zodResolver(completeSchema),
    defaultValues: {
      completion_note: '',
      time_minutes: undefined,
      costs: [],
      asset_status: '',
      follow_up_title: '',
    },
  });
  const costs = useFieldArray({ control: form.control, name: 'costs' });
  const handle = useFieldErrors<CompleteInput>((n, e) => form.setError(n, e), setBanner);
  const submit = form.handleSubmit(async (values) => {
    setBanner(null);
    const data = completeSchema.parse(values);
    try {
      const result = await api<WorkOrderDetail>(`/work-orders/${wo.id}/complete`, {
        method: 'POST',
        body: data,
      });
      toast.success(
        `#${wo.number} completed`,
        result.follow_up
          ? `Follow-up #${result.follow_up.number} created`
          : result.next_occurrence
            ? `Next occurrence #${result.next_occurrence.number} scheduled`
            : undefined,
      );
      onDone(result);
    } catch (err) {
      handle(err);
    }
  });
  const e = form.formState.errors;
  return (
    <Dialog
      open={open}
      onClose={onClose}
      title={`Complete #${wo.number}`}
      description={
        wo.children.length > 0 && wo.child_progress.done < wo.child_progress.total
          ? `${wo.child_progress.total - wo.child_progress.done} sub-work order(s) are still open; they stay open.`
          : undefined
      }
      footer={
        <>
          <Button variant="ghost" onClick={onClose}>
            Cancel
          </Button>
          <Button
            variant="success"
            onClick={submit}
            loading={form.formState.isSubmitting}
            data-testid="complete-submit"
          >
            Mark as done
          </Button>
        </>
      }
    >
      <form className="stack" style={{ gap: 16 }} onSubmit={submit} noValidate>
        {banner && (
          <div className="form-banner form-banner-error" role="alert">
            {banner}
          </div>
        )}
        <TextArea
          label="Completion note"
          placeholder="What was done, what was found, anything the next person should know."
          error={e.completion_note?.message}
          {...form.register('completion_note')}
          data-autofocus
          data-testid="complete-note"
          rows={3}
        />
        <div className="field-row">
          <TextField
            label="Time spent (minutes)"
            type="number"
            min={0}
            step={5}
            error={e.time_minutes?.message}
            {...form.register('time_minutes', { valueAsNumber: true })}
            data-testid="complete-minutes"
          />
          {wo.primary_asset && (
            <SelectField
              label={`${wo.primary_asset.name} status after work`}
              options={ASSET_STATUSES.map((s) => ({ value: s, label: STATUS_LABEL[s] }))}
              placeholder="Leave unchanged"
              {...form.register('asset_status')}
            />
          )}
        </div>
        <div className="field">
          <span className="field-label">Costs</span>
          {costs.fields.map((f, i) => (
            <div key={f.id} className="row" style={{ alignItems: 'flex-start' }}>
              <SelectField
                aria-label="Cost type"
                options={[
                  { value: 'part', label: 'Part' },
                  { value: 'labor', label: 'Labor' },
                  { value: 'vendor', label: 'Vendor' },
                  { value: 'other', label: 'Other' },
                ]}
                {...form.register(`costs.${i}.type`)}
                wrapClassName="cost-type"
              />
              <TextField
                aria-label="Amount"
                type="number"
                min={0}
                step="0.01"
                placeholder="0.00"
                error={e.costs?.[i]?.amount?.message}
                {...form.register(`costs.${i}.amount`, { valueAsNumber: true })}
              />
              <TextField
                aria-label="Description"
                placeholder="Description"
                {...form.register(`costs.${i}.description`)}
              />
              <Button
                variant="ghost"
                onClick={() => costs.remove(i)}
                aria-label="Remove cost"
                icon={<Trash2 />}
              />
            </div>
          ))}
          <div>
            <Button
              size="sm"
              icon={<Plus />}
              onClick={() => costs.append({ type: 'part', amount: 0, description: '' })}
            >
              Add cost
            </Button>
          </div>
        </div>
        <TextField
          label="Create a follow-up work order"
          placeholder="Optional: title for the follow-up"
          error={e.follow_up_title?.message}
          {...form.register('follow_up_title')}
          hint="Copies project, team, asset and location from this work order."
        />
      </form>
    </Dialog>
  );
}

const noteSchema = z.object({ note: z.string().trim().max(500).optional().or(z.literal('')) });
type NoteInput = z.infer<typeof noteSchema>;

export function NoteDialog({
  wo,
  open,
  onClose,
  onDone,
  action,
  title,
  label,
  confirmLabel,
  danger,
  required,
}: {
  wo: WorkOrderDetail;
  open: boolean;
  onClose: () => void;
  onDone: Done;
  action: 'hold' | 'cancel' | 'start' | 'resume';
  title: string;
  label: string;
  confirmLabel: string;
  danger?: boolean;
  required?: boolean;
}) {
  const toast = useToast();
  const [banner, setBanner] = useState<string | null>(null);
  const form = useForm<NoteInput>({ resolver: zodResolver(noteSchema), defaultValues: { note: '' } });
  const submit = form.handleSubmit(async (values) => {
    if (required && !values.note?.trim()) {
      form.setError('note', { message: 'Please give a reason.' });
      return;
    }
    setBanner(null);
    try {
      const result = await api<WorkOrderDetail>(`/work-orders/${wo.id}/${action}`, {
        method: 'POST',
        body: { note: values.note?.trim() || null },
      });
      toast.success(
        `#${wo.number} ${action === 'cancel' ? 'canceled' : action === 'hold' ? 'put on hold' : action === 'resume' ? 'resumed' : 'started'}`,
      );
      onDone(result);
    } catch (err) {
      setBanner(err instanceof ApiError ? err.message : 'Something went wrong.');
    }
  });
  return (
    <Dialog
      open={open}
      onClose={onClose}
      title={title}
      footer={
        <>
          <Button variant="ghost" onClick={onClose}>
            Back
          </Button>
          <Button
            variant={danger ? 'danger' : 'primary'}
            onClick={submit}
            loading={form.formState.isSubmitting}
            data-testid="note-submit"
          >
            {confirmLabel}
          </Button>
        </>
      }
    >
      <form className="stack" onSubmit={submit} noValidate>
        {banner && (
          <div className="form-banner form-banner-error" role="alert">
            {banner}
          </div>
        )}
        <TextArea
          label={label}
          required={required}
          error={form.formState.errors.note?.message}
          {...form.register('note')}
          data-autofocus
          rows={3}
        />
      </form>
    </Dialog>
  );
}

export function AssignDialog({
  wo,
  open,
  onClose,
  onDone,
}: {
  wo: WorkOrderDetail;
  open: boolean;
  onClose: () => void;
  onDone: Done;
}) {
  const toast = useToast();
  const lookups = useLookups();
  const [users, setUsers] = useState<number[]>(wo.assignees.map((u) => u.id));
  const [teams, setTeams] = useState<string[]>(wo.assignee_teams.map((t) => t.id));
  const [busy, setBusy] = useState(false);
  const [banner, setBanner] = useState<string | null>(null);
  return (
    <Dialog
      open={open}
      onClose={onClose}
      title="Assign"
      description="Assignees are notified and the work order appears in their To Do list."
      footer={
        <>
          <Button variant="ghost" onClick={onClose}>
            Cancel
          </Button>
          <Button
            variant="primary"
            loading={busy}
            data-testid="assign-submit"
            onClick={async () => {
              setBusy(true);
              setBanner(null);
              try {
                const result = await api<WorkOrderDetail>(`/work-orders/${wo.id}/assignees`, {
                  method: 'PUT',
                  body: { user_ids: users, team_ids: teams },
                });
                toast.success('Assignees updated');
                onDone(result);
              } catch (err) {
                setBanner(err instanceof ApiError ? err.message : 'Something went wrong.');
              } finally {
                setBusy(false);
              }
            }}
          >
            Save assignees
          </Button>
        </>
      }
    >
      <div className="stack" style={{ gap: 16 }}>
        {banner && (
          <div className="form-banner form-banner-error" role="alert">
            {banner}
          </div>
        )}
        <Picker
          label="People"
          options={userOptions(lookups.data)}
          value={users}
          onChange={setUsers}
          placeholder="Choose people"
          id="assign-users"
        />
        <Picker
          label="Teams"
          options={teamOptions(lookups.data, wo.project?.id || null)}
          value={teams}
          onChange={setTeams}
          placeholder="Choose teams"
        />
      </div>
    </Dialog>
  );
}

export function WatchersDialog({
  wo,
  open,
  onClose,
  onDone,
}: {
  wo: WorkOrderDetail;
  open: boolean;
  onClose: () => void;
  onDone: Done;
}) {
  const toast = useToast();
  const lookups = useLookups();
  const [users, setUsers] = useState<number[]>(wo.watchers.map((u) => u.id));
  const [busy, setBusy] = useState(false);
  return (
    <Dialog
      open={open}
      onClose={onClose}
      title="Watchers"
      description="Watchers get notified about status changes and comments without being assigned."
      footer={
        <>
          <Button variant="ghost" onClick={onClose}>
            Cancel
          </Button>
          <Button
            variant="primary"
            loading={busy}
            onClick={async () => {
              setBusy(true);
              try {
                const result = await api<WorkOrderDetail>(`/work-orders/${wo.id}/watchers`, {
                  method: 'PUT',
                  body: { user_ids: users },
                });
                toast.success('Watchers updated');
                onDone(result);
              } catch (err) {
                toast.error('Could not update watchers', err instanceof ApiError ? err.message : undefined);
              } finally {
                setBusy(false);
              }
            }}
          >
            Save watchers
          </Button>
        </>
      }
    >
      <Picker
        label="People"
        options={userOptions(lookups.data)}
        value={users}
        onChange={setUsers}
        placeholder="Choose people"
      />
    </Dialog>
  );
}

const timeSchema = z.object({
  minutes: z.number().int().min(1, 'At least one minute').max(100000),
  note: z.string().trim().max(500).optional().or(z.literal('')),
});
type TimeInput = z.infer<typeof timeSchema>;

export function LogTimeDialog({
  wo,
  open,
  onClose,
  onDone,
}: {
  wo: WorkOrderDetail;
  open: boolean;
  onClose: () => void;
  onDone: () => void;
}) {
  const toast = useToast();
  const form = useForm<TimeInput>({
    resolver: zodResolver(timeSchema),
    defaultValues: { minutes: 30, note: '' },
  });
  const [banner, setBanner] = useState<string | null>(null);
  const submit = form.handleSubmit(async (values) => {
    setBanner(null);
    try {
      await api(`/work-orders/${wo.id}/time-entries`, {
        method: 'POST',
        body: { minutes: values.minutes, note: values.note?.trim() || null },
      });
      toast.success('Time logged');
      onDone();
    } catch (err) {
      setBanner(err instanceof ApiError ? err.message : 'Something went wrong.');
    }
  });
  return (
    <Dialog
      open={open}
      onClose={onClose}
      title="Log time"
      footer={
        <>
          <Button variant="ghost" onClick={onClose}>
            Cancel
          </Button>
          <Button
            variant="primary"
            onClick={submit}
            loading={form.formState.isSubmitting}
            data-testid="time-submit"
          >
            Log time
          </Button>
        </>
      }
    >
      <form className="stack" onSubmit={submit} noValidate>
        {banner && (
          <div className="form-banner form-banner-error" role="alert">
            {banner}
          </div>
        )}
        <TextField
          label="Minutes"
          type="number"
          min={1}
          step={5}
          required
          error={form.formState.errors.minutes?.message}
          {...form.register('minutes', { valueAsNumber: true })}
          data-autofocus
          data-testid="time-minutes"
        />
        <TextField label="Note" placeholder="Optional" {...form.register('note')} />
      </form>
    </Dialog>
  );
}

const costSchema = z.object({
  type: z.enum(['part', 'labor', 'vendor', 'other']),
  amount: z.number().min(0, 'Amount cannot be negative'),
  description: z.string().trim().max(300).optional().or(z.literal('')),
});
type CostInputForm = z.infer<typeof costSchema>;

export function AddCostDialog({
  wo,
  open,
  onClose,
  onDone,
}: {
  wo: WorkOrderDetail;
  open: boolean;
  onClose: () => void;
  onDone: () => void;
}) {
  const toast = useToast();
  const form = useForm<CostInputForm>({
    resolver: zodResolver(costSchema),
    defaultValues: { type: 'part', amount: 0, description: '' },
  });
  const [banner, setBanner] = useState<string | null>(null);
  const submit = form.handleSubmit(async (values) => {
    setBanner(null);
    try {
      await api(`/work-orders/${wo.id}/cost-entries`, {
        method: 'POST',
        body: { type: values.type, amount: values.amount, description: values.description?.trim() || null },
      });
      toast.success('Cost added');
      onDone();
    } catch (err) {
      setBanner(err instanceof ApiError ? err.message : 'Something went wrong.');
    }
  });
  return (
    <Dialog
      open={open}
      onClose={onClose}
      title="Add cost"
      footer={
        <>
          <Button variant="ghost" onClick={onClose}>
            Cancel
          </Button>
          <Button variant="primary" onClick={submit} loading={form.formState.isSubmitting}>
            Add cost
          </Button>
        </>
      }
    >
      <form className="stack" onSubmit={submit} noValidate>
        {banner && (
          <div className="form-banner form-banner-error" role="alert">
            {banner}
          </div>
        )}
        <div className="field-row">
          <SelectField
            label="Type"
            options={[
              { value: 'part', label: 'Part' },
              { value: 'labor', label: 'Labor' },
              { value: 'vendor', label: 'Vendor' },
              { value: 'other', label: 'Other' },
            ]}
            {...form.register('type')}
          />
          <TextField
            label="Amount (USD)"
            type="number"
            min={0}
            step="0.01"
            required
            error={form.formState.errors.amount?.message}
            {...form.register('amount', { valueAsNumber: true })}
            data-autofocus
          />
        </div>
        <TextField label="Description" placeholder="e.g. 4x M4 bolts" {...form.register('description')} />
      </form>
    </Dialog>
  );
}

type SubInput = z.input<typeof subWorkOrderSchema>;

export function SubWorkOrderDialog({
  parent,
  open,
  onClose,
  onDone,
}: {
  parent: WorkOrderDetail;
  open: boolean;
  onClose: () => void;
  onDone: (child: WorkOrderDetail) => void;
}) {
  const toast = useToast();
  const lookups = useLookups();
  const [banner, setBanner] = useState<string | null>(null);
  const form = useForm<SubInput>({
    resolver: zodResolver(subWorkOrderSchema),
    defaultValues: {
      title: '',
      description: '',
      assignee_user_ids: [],
      due_at: '',
      priority: parent.priority === 'CRITICAL' ? 'HIGH' : parent.priority,
      estimated_minutes: undefined,
    },
  });
  const handle = useFieldErrors<SubInput>((n, e) => form.setError(n, e), setBanner);
  const submit = form.handleSubmit(async (values) => {
    setBanner(null);
    const data = subWorkOrderSchema.parse(values);
    try {
      const child = await api<WorkOrderDetail>(`/work-orders/${parent.id}/sub-work-orders`, {
        method: 'POST',
        body: data,
      });
      toast.success(`Sub-work order #${child.number} created`);
      onDone(child);
    } catch (err) {
      handle(err);
    }
  });
  const e = form.formState.errors;
  return (
    <Dialog
      open={open}
      onClose={onClose}
      title={`Add sub-work order to #${parent.number}`}
      description="Sub-work orders inherit the project, team, asset and location of their parent."
      footer={
        <>
          <Button variant="ghost" onClick={onClose}>
            Cancel
          </Button>
          <Button
            variant="primary"
            onClick={submit}
            loading={form.formState.isSubmitting}
            data-testid="sub-submit"
          >
            Create sub-work order
          </Button>
        </>
      }
    >
      <form className="stack" style={{ gap: 16 }} onSubmit={submit} noValidate>
        {banner && (
          <div className="form-banner form-banner-error" role="alert">
            {banner}
          </div>
        )}
        <TextField
          label="Title"
          required
          error={e.title?.message}
          {...form.register('title')}
          data-autofocus
          data-testid="sub-title"
        />
        <TextArea label="Description" rows={2} {...form.register('description')} />
        <Controller
          control={form.control}
          name="assignee_user_ids"
          render={({ field }) => (
            <Picker
              label="Assignees"
              options={userOptions(lookups.data)}
              value={field.value || []}
              onChange={field.onChange}
              placeholder="Assign people"
            />
          )}
        />
        <div className="field-row">
          <TextField
            label="Due"
            type="datetime-local"
            error={e.due_at?.message}
            {...form.register('due_at')}
          />
          <SelectField
            label="Priority"
            options={WO_PRIORITIES.map((p) => ({ value: p, label: humanize(p) }))}
            {...form.register('priority')}
          />
        </div>
        <TextField
          label="Estimated minutes"
          type="number"
          min={0}
          step={5}
          error={e.estimated_minutes?.message}
          {...form.register('estimated_minutes', { valueAsNumber: true })}
        />
      </form>
    </Dialog>
  );
}
