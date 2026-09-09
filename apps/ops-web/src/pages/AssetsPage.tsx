import { useMemo, useState, type ReactNode } from 'react';
import { Link } from 'react-router-dom';
import { Controller, useForm } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import { useQuery } from '@tanstack/react-query';
import { Activity, ArrowLeft, Boxes, Pencil, Plus, Power, X } from 'lucide-react';
import { api, ApiError } from '@/api/client';
import { keys, useAsset, useAssetHistory, useAssets, useInvalidate, useLookups } from '@/api/hooks';
import { useCurrentSession } from '@/auth/SessionProvider';
import { assetSchema, type AssetInput } from '@/contracts/schemas';
import type { Asset, AssetStatus } from '@/contracts/types';
import { useListParams } from '@/lib/listParams';
import { useIsCompact } from '@/lib/useMediaQuery';
import { fmtDate, fmtMoney, humanize, STATUS_LABEL } from '@/lib/format';
import {
  ASSET_STATUSES,
  locationFilterOptions,
  locationOptions,
  projectFilterOptions,
  projectOptions,
  teamFilterOptions,
  teamOptions,
  userOptions,
} from '@/lib/options';
import { ActivityTimeline } from '@/ui/ActivityTimeline';
import { AttachmentList, AttachmentUploader } from '@/ui/Attachments';
import {
  Badge,
  Button,
  Dialog,
  EmptyState,
  ErrorState,
  FilterChip,
  IconButton,
  PageHeader,
  Picker,
  SearchField,
  SelectField,
  SideSheet,
  Skeleton,
  TextArea,
  TextField,
  useToast,
} from '@/ui';
import { cn } from '@/lib/cn';

const STATUS_TONE: Record<string, string> = {
  ONLINE: 'success',
  OFFLINE_PLANNED: 'warning',
  OFFLINE_UNPLANNED: 'critical',
  DO_NOT_TRACK: 'neutral',
  RETIRED: 'neutral',
};
const CRIT_TONE: Record<string, string> = {
  low: 'neutral',
  medium: 'info',
  high: 'warning',
  critical: 'critical',
};

export function AssetsPage() {
  const { can, session } = useCurrentSession();
  const lookups = useLookups();
  const list = useListParams(['asset', 'new']);
  const compact = useIsCompact();
  const selectedId = list.state.extras.asset || null;
  const creating = list.state.extras.new === '1';
  const params = useMemo(() => list.apiParams, [list.apiParams]);
  const query = useAssets(params);
  const types = useQuery<{ id: string; name: string; color: string; icon: string }[]>({
    queryKey: ['asset-types'],
    queryFn: () => api('/asset-types'),
    staleTime: 300_000,
  });
  const items = query.data?.items || [];
  const filterValue = (key: string) => list.state.filters[key] || [];
  const setFilter = (key: string) => (values: string[]) => list.update({ filters: { [key]: values } });
  const canCreate = can('asset.manage');
  const select = (id: string | null) => list.update({ extras: { asset: id } }, false);

  return (
    <div className="page">
      <PageHeader
        title="Assets"
        subtitle="The rover, its subsystems and the shop equipment you maintain."
        testId="assets-page"
        actions={
          canCreate && (
            <Button
              variant="primary"
              icon={<Plus />}
              onClick={() => list.update({ extras: { new: '1' } })}
              data-testid="new-asset"
            >
              New asset
            </Button>
          )
        }
      />
      <div className="page-toolbar">
        <SearchField
          value={list.state.q}
          onChange={(q) => list.update({ q })}
          placeholder="Search name or code"
        />
        <div className="filter-bar">
          <FilterChip
            label="Status"
            options={ASSET_STATUSES.map((s) => ({ value: s, label: STATUS_LABEL[s] }))}
            value={filterValue('status')}
            onChange={setFilter('status')}
          />
          <FilterChip
            label="Criticality"
            options={['low', 'medium', 'high', 'critical'].map((c) => ({ value: c, label: humanize(c) }))}
            value={filterValue('criticality')}
            onChange={setFilter('criticality')}
          />
          <FilterChip
            label="Project"
            options={projectFilterOptions(lookups.data)}
            value={filterValue('project')}
            onChange={setFilter('project')}
          />
          <FilterChip
            label="Location"
            options={locationFilterOptions(lookups.data)}
            value={filterValue('location')}
            onChange={setFilter('location')}
          />
          <FilterChip
            label="Team"
            options={teamFilterOptions(lookups.data)}
            value={filterValue('team')}
            onChange={setFilter('team')}
          />
          <FilterChip
            label="Type"
            options={(types.data || []).map((t) => ({ value: t.id, label: t.name, color: t.color }))}
            value={filterValue('type')}
            onChange={setFilter('type')}
          />
          {list.activeFilterCount > 0 && (
            <button type="button" className="filter-chip-clear" onClick={list.clearFilters}>
              Clear all
            </button>
          )}
        </div>
      </div>
      <div className="page-body">
        <div className={cn('master-detail', !selectedId && 'no-detail', selectedId && 'has-detail')}>
          <div className="master">
            {query.isPending ? (
              <Skeleton lines={6} />
            ) : query.error ? (
              <ErrorState error={query.error} onRetry={() => query.refetch()} />
            ) : items.length === 0 ? (
              <div className="card">
                <EmptyState
                  icon={<Boxes />}
                  title={list.activeFilterCount ? 'No assets match' : 'No assets yet'}
                  body="Register the rover and the equipment you maintain so work orders build a history for each."
                  actions={
                    canCreate && !list.activeFilterCount ? (
                      <Button
                        variant="primary"
                        icon={<Plus />}
                        onClick={() => list.update({ extras: { new: '1' } })}
                      >
                        New asset
                      </Button>
                    ) : undefined
                  }
                />
              </div>
            ) : (
              <div className="table-wrap">
                <table className="table" data-testid="assets-table">
                  <thead>
                    <tr>
                      <th>Asset</th>
                      <th>Status</th>
                      <th>Criticality</th>
                      <th>Location</th>
                      <th>Team</th>
                      <th className="num">Open WOs</th>
                    </tr>
                  </thead>
                  <tbody>
                    {items.map((a) => (
                      <tr
                        key={a.id}
                        className={cn('clickable', a.id === selectedId && 'selected')}
                        onClick={() => select(a.id)}
                        tabIndex={0}
                        onKeyDown={(e) => e.key === 'Enter' && select(a.id)}
                        data-testid="asset-row"
                      >
                        <td>
                          <span style={{ fontWeight: 600 }}>{a.name}</span>
                          <div className="text-caption text-muted">
                            {a.code}
                            {a.parent ? ` · in ${a.parent.name}` : ''}
                            {a.child_count ? ` · ${a.child_count} sub-assets` : ''}
                          </div>
                        </td>
                        <td>
                          <Badge tone={STATUS_TONE[a.status]}>{STATUS_LABEL[a.status]}</Badge>
                        </td>
                        <td>
                          <Badge tone={CRIT_TONE[a.criticality]}>{humanize(a.criticality)}</Badge>
                        </td>
                        <td>{a.location?.name || '—'}</td>
                        <td>{a.responsible_team?.name || '—'}</td>
                        <td className="num">{a.open_work_orders}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
          {selectedId && (
            <div className="detail">
              <AssetDetail
                id={selectedId}
                onClose={() => select(null)}
                onNavigate={select}
                showBack={compact}
                canManage={can('asset.manage')}
                currentUserId={session.user.id}
              />
            </div>
          )}
        </div>
      </div>
      {creating && (
        <AssetForm
          onClose={() => list.update({ extras: { new: null } })}
          onSaved={(a) => {
            list.update({ extras: { new: null, asset: a.id } });
          }}
        />
      )}
    </div>
  );
}

function AssetDetail({
  id,
  onClose,
  onNavigate,
  showBack,
  canManage,
  currentUserId,
}: {
  id: string;
  onClose: () => void;
  onNavigate: (id: string) => void;
  showBack: boolean;
  canManage: boolean;
  currentUserId: number;
}) {
  const asset = useAsset(id);
  const history = useAssetHistory(id);
  const invalidate = useInvalidate();
  const [editing, setEditing] = useState(false);
  const [changingStatus, setChangingStatus] = useState(false);
  const { session } = useCurrentSession();
  if (asset.isPending)
    return (
      <div className="detail-panel">
        <Skeleton lines={8} />
      </div>
    );
  if (asset.error || !asset.data)
    return (
      <div className="detail-panel" style={{ padding: 16 }}>
        <ErrorState error={asset.error} onRetry={() => asset.refetch()} />
      </div>
    );
  const a = asset.data;
  const editable = canManage || a.permissions?.edit;
  return (
    <article className="detail-panel" aria-label={a.name} data-testid="asset-detail">
      <div className="detail-head">
        <div className="row-between">
          <div className="row">
            {showBack && (
              <IconButton label="Back to list" onClick={onClose}>
                <ArrowLeft />
              </IconButton>
            )}
            <span className="mono text-muted">{a.code}</span>
            <Badge tone={STATUS_TONE[a.status]}>{STATUS_LABEL[a.status]}</Badge>
            <Badge tone={CRIT_TONE[a.criticality]}>{humanize(a.criticality)}</Badge>
          </div>
          {!showBack && (
            <IconButton label="Close details" onClick={onClose}>
              <X />
            </IconButton>
          )}
        </div>
        <h2 className="detail-title">{a.name}</h2>
        {a.parent && (
          <div className="text-label text-muted">
            Part of{' '}
            <button type="button" className="link-button" onClick={() => onNavigate(a.parent!.id)}>
              {a.parent.name}
            </button>
          </div>
        )}
        <div className="detail-actions">
          {editable && (
            <Button
              variant="primary"
              icon={<Power />}
              onClick={() => setChangingStatus(true)}
              data-testid="asset-status"
            >
              Change status
            </Button>
          )}
          {editable && (
            <Button icon={<Pencil />} onClick={() => setEditing(true)}>
              Edit
            </Button>
          )}
          <Link to={`/work-orders?filter[asset]=${a.id}`} className="btn btn-secondary">
            Work orders ({a.open_work_orders} open)
          </Link>
        </div>
      </div>
      <section className="detail-section">
        <div className="detail-grid">
          <Field label="Project">
            {a.project ? <Link to={`/projects/${a.project.id}`}>{a.project.name}</Link> : '—'}
          </Field>
          <Field label="Location">{a.location?.name || '—'}</Field>
          <Field label="Responsible team">{a.responsible_team?.name || '—'}</Field>
          <Field label="Owner">{a.owner?.name || '—'}</Field>
          <Field label="Manufacturer / model">
            {[a.manufacturer, a.model].filter(Boolean).join(' ') || '—'}
          </Field>
          <Field label="Serial">{a.serial_number || '—'}</Field>
          <Field label="Purchased">
            {a.purchase_date ? `${fmtDate(a.purchase_date)} · ${fmtMoney(a.purchase_cost)}` : '—'}
          </Field>
          <Field label="Warranty">{a.warranty_end ? fmtDate(a.warranty_end) : '—'}</Field>
          <Field label="Types">
            <div className="row" style={{ flexWrap: 'wrap' }}>
              {a.asset_types.map((t) => (
                <span key={t.id} className="chip">
                  <span className="chip-dot" style={{ background: t.color }} />
                  {t.name}
                </span>
              ))}
              {a.asset_types.length === 0 && '—'}
            </div>
          </Field>
        </div>
        {a.description && (
          <div style={{ marginTop: 16 }}>
            <div className="detail-field-label">Description</div>
            <div className="description-block">{a.description}</div>
          </div>
        )}
      </section>
      {(a.children || []).length > 0 && (
        <section className="detail-section">
          <div className="detail-section-title">
            <span>
              Sub-assets <span className="count">{a.children!.length}</span>
            </span>
          </div>
          {a.children!.map((c) => (
            <button
              key={c.id}
              type="button"
              className="list-row"
              style={{ borderLeft: 'none', paddingLeft: 8, paddingRight: 8 }}
              onClick={() => onNavigate(c.id)}
            >
              <div className="list-row-main">
                <div className="list-row-title">{c.name}</div>
                <div className="list-row-meta">{c.code}</div>
              </div>
              <div className="list-row-side">
                <Badge tone={STATUS_TONE[c.status]}>{STATUS_LABEL[c.status]}</Badge>
              </div>
            </button>
          ))}
        </section>
      )}
      <section className="detail-section">
        <div className="detail-section-title">
          <span>
            Files <span className="count">{(a.files || []).length}</span>
          </span>
        </div>
        <div className="stack">
          <AttachmentList
            files={a.files || []}
            currentUser={session.user}
            canDelete={(f) => !!editable || f.uploaded_by?.id === currentUserId}
            onChanged={() => invalidate(keys.asset(id))}
          />
          {(editable || session.permissions.includes('file.upload')) && (
            <AttachmentUploader
              entityType="asset"
              entityId={a.id}
              onUploaded={() => invalidate(keys.asset(id))}
              compact
            />
          )}
        </div>
      </section>
      <section className="detail-section">
        <div className="detail-section-title">
          <span className="row">
            <Activity size={16} /> Status history
          </span>
        </div>
        {history.isPending ? (
          <Skeleton lines={3} />
        ) : (
          <ActivityTimeline
            items={(history.data || []).map((h) => ({
              id: h.id,
              title: `${h.from_status ? `${STATUS_LABEL[h.from_status]} → ` : ''}${STATUS_LABEL[h.to_status]}`,
              meta: `${h.changed_by?.name || 'System'} · ${fmtDate(h.started_at, true)}${h.ended_at ? ` → ${fmtDate(h.ended_at, true)}` : ''}`,
              note:
                [h.downtime_type ? `${humanize(h.downtime_type)} downtime` : null, h.downtime_reason, h.note]
                  .filter(Boolean)
                  .join(' · ') || null,
              tone:
                h.to_status === 'ONLINE'
                  ? 'success'
                  : h.to_status === 'OFFLINE_UNPLANNED'
                    ? 'danger'
                    : 'default',
            }))}
            emptyText="No status changes recorded."
          />
        )}
      </section>
      {editing && <AssetForm asset={a} onClose={() => setEditing(false)} onSaved={() => setEditing(false)} />}
      {changingStatus && <StatusDialog asset={a} onClose={() => setChangingStatus(false)} />}
    </article>
  );
}

function Field({ label, children }: { label: string; children: ReactNode }) {
  return (
    <div>
      <div className="detail-field-label">{label}</div>
      <div className="detail-field-value">{children}</div>
    </div>
  );
}

function StatusDialog({ asset, onClose }: { asset: Asset; onClose: () => void }) {
  const invalidate = useInvalidate();
  const toast = useToast();
  const [status, setStatus] = useState<AssetStatus>(asset.status);
  const [downtimeType, setDowntimeType] = useState<'planned' | 'unplanned' | ''>('');
  const [reason, setReason] = useState('');
  const [note, setNote] = useState('');
  const [busy, setBusy] = useState(false);
  const offline = status === 'OFFLINE_PLANNED' || status === 'OFFLINE_UNPLANNED';
  return (
    <Dialog
      open
      onClose={onClose}
      title={`Change status of ${asset.name}`}
      footer={
        <>
          <Button variant="ghost" onClick={onClose}>
            Cancel
          </Button>
          <Button
            variant="primary"
            loading={busy}
            data-testid="asset-status-submit"
            onClick={async () => {
              setBusy(true);
              try {
                await api(`/assets/${asset.id}/status`, {
                  method: 'POST',
                  body: {
                    status,
                    downtime_type: offline
                      ? downtimeType || (status === 'OFFLINE_PLANNED' ? 'planned' : 'unplanned')
                      : null,
                    downtime_reason: reason || null,
                    note: note || null,
                  },
                });
                invalidate(keys.asset(asset.id), keys.assetHistory(asset.id), ['assets'], keys.lookups);
                toast.success('Status updated');
                onClose();
              } catch (err) {
                toast.error('Could not change status', err instanceof ApiError ? err.message : undefined);
              } finally {
                setBusy(false);
              }
            }}
          >
            Update status
          </Button>
        </>
      }
    >
      <div className="stack" style={{ gap: 16 }}>
        <SelectField
          label="Status"
          options={ASSET_STATUSES.map((s) => ({ value: s, label: STATUS_LABEL[s] }))}
          value={status}
          onChange={(e) => setStatus(e.target.value as AssetStatus)}
          data-autofocus
        />
        {offline && (
          <div className="field-row">
            <SelectField
              label="Downtime type"
              options={[
                { value: 'planned', label: 'Planned' },
                { value: 'unplanned', label: 'Unplanned' },
              ]}
              value={downtimeType || (status === 'OFFLINE_PLANNED' ? 'planned' : 'unplanned')}
              onChange={(e) => setDowntimeType(e.target.value as 'planned' | 'unplanned')}
            />
            <TextField
              label="Reason"
              placeholder="e.g. Motor controller failure"
              value={reason}
              onChange={(e) => setReason(e.target.value)}
              maxLength={160}
            />
          </div>
        )}
        <TextArea
          label="Note"
          rows={2}
          value={note}
          onChange={(e) => setNote(e.target.value)}
          maxLength={500}
        />
      </div>
    </Dialog>
  );
}

function AssetForm({
  asset,
  onClose,
  onSaved,
}: {
  asset?: Asset;
  onClose: () => void;
  onSaved: (asset: Asset) => void;
}) {
  const lookups = useLookups();
  const invalidate = useInvalidate();
  const toast = useToast();
  const types = useQuery<{ id: string; name: string; color: string; icon: string }[]>({
    queryKey: ['asset-types'],
    queryFn: () => api('/asset-types'),
    staleTime: 300_000,
  });
  const [banner, setBanner] = useState<string | null>(null);
  const form = useForm<AssetInput>({
    resolver: zodResolver(assetSchema),
    defaultValues: {
      name: asset?.name || '',
      code: asset?.code || '',
      description: asset?.description || '',
      parent_asset_id: asset?.parent_asset_id || '',
      project_id: asset?.project_id || '',
      location_id: asset?.location_id || '',
      responsible_team_id: asset?.responsible_team_id || '',
      owner_user_id: asset?.owner?.id ?? null,
      manufacturer: asset?.manufacturer || '',
      model: asset?.model || '',
      serial_number: asset?.serial_number || '',
      purchase_date: asset?.purchase_date || '',
      purchase_cost: asset?.purchase_cost ?? undefined,
      warranty_end: asset?.warranty_end || '',
      criticality: asset?.criticality || 'medium',
      status: asset?.status || 'ONLINE',
      asset_type_ids: asset?.asset_types.map((t) => t.id) || [],
    },
  });
  const projectId = form.watch('project_id');
  const submit = form.handleSubmit(async (values) => {
    setBanner(null);
    const data = assetSchema.parse(values);
    try {
      let saved: Asset;
      if (asset) {
        const body: Record<string, unknown> = {};
        const dirty = form.formState.dirtyFields as Record<string, unknown>;
        for (const key of Object.keys(data) as (keyof AssetInput)[])
          if (dirty[key] && key !== 'status') body[key] = data[key] ?? null;
        saved = Object.keys(body).length
          ? await api<Asset>(`/assets/${asset.id}`, { method: 'PATCH', body })
          : asset;
        toast.success('Asset updated');
        invalidate(keys.asset(asset.id));
      } else {
        saved = await api<Asset>('/assets', { method: 'POST', body: data });
        toast.success('Asset created');
      }
      invalidate(['assets'], keys.lookups, keys.setup);
      onSaved(saved);
    } catch (err) {
      if (err instanceof ApiError) {
        const fields = err.fieldErrors();
        Object.entries(fields).forEach(([f, m]) => form.setError(f as keyof AssetInput, { message: m }));
        if (!Object.keys(fields).length) setBanner(err.message);
      }
    }
  });
  const e = form.formState.errors;
  return (
    <SideSheet
      open
      onClose={onClose}
      title={asset ? `Edit ${asset.name}` : 'New asset'}
      testId="asset-form"
      footer={
        <>
          <span className="grow" />
          <Button variant="ghost" onClick={onClose}>
            Cancel
          </Button>
          <Button
            variant="primary"
            onClick={submit}
            loading={form.formState.isSubmitting}
            data-testid="asset-submit"
          >
            {asset ? 'Save changes' : 'Create asset'}
          </Button>
        </>
      }
    >
      <form className="stack" style={{ gap: 18 }} onSubmit={submit} noValidate>
        {banner && (
          <div className="form-banner form-banner-error" role="alert">
            {banner}
          </div>
        )}
        <div className="field-row">
          <TextField
            label="Name"
            required
            error={e.name?.message}
            {...form.register('name')}
            data-autofocus
            data-testid="asset-name"
          />
          <TextField
            label="Code"
            placeholder="Auto from name"
            error={e.code?.message}
            {...form.register('code')}
          />
        </div>
        <TextArea label="Description" rows={2} {...form.register('description')} />
        <Controller
          control={form.control}
          name="asset_type_ids"
          render={({ field }) => (
            <Picker
              label="Types"
              options={(types.data || []).map((t) => ({ value: t.id, label: t.name, color: t.color }))}
              value={field.value || []}
              onChange={field.onChange}
              placeholder="e.g. 3D Printer"
            />
          )}
        />
        <div className="form-section">
          <div className="form-section-title">Where &amp; who</div>
          <div className="field-row">
            <Controller
              control={form.control}
              name="project_id"
              render={({ field }) => (
                <Picker
                  label="Project"
                  multiple={false}
                  options={projectOptions(lookups.data)}
                  value={field.value ? [field.value] : []}
                  onChange={(v) => field.onChange(v[0] || '')}
                  placeholder="Shared equipment"
                />
              )}
            />
            <Controller
              control={form.control}
              name="parent_asset_id"
              render={({ field }) => (
                <Picker
                  label="Part of"
                  multiple={false}
                  options={(lookups.data?.assets || [])
                    .filter((x) => x.id !== asset?.id)
                    .map((x) => ({ value: x.id, label: x.name, hint: x.code || undefined }))}
                  value={field.value ? [field.value] : []}
                  onChange={(v) => field.onChange(v[0] || '')}
                  placeholder="Top-level asset"
                />
              )}
            />
          </div>
          <div className="field-row">
            <Controller
              control={form.control}
              name="location_id"
              render={({ field }) => (
                <Picker
                  label="Location"
                  multiple={false}
                  options={locationOptions(lookups.data)}
                  value={field.value ? [field.value] : []}
                  onChange={(v) => field.onChange(v[0] || '')}
                  placeholder="Default"
                />
              )}
            />
            <Controller
              control={form.control}
              name="responsible_team_id"
              render={({ field }) => (
                <Picker
                  label="Responsible team"
                  multiple={false}
                  options={teamOptions(lookups.data, projectId || null)}
                  value={field.value ? [field.value] : []}
                  onChange={(v) => field.onChange(v[0] || '')}
                  placeholder="None"
                />
              )}
            />
          </div>
          <Controller
            control={form.control}
            name="owner_user_id"
            render={({ field }) => (
              <Picker
                label="Owner"
                multiple={false}
                options={userOptions(lookups.data)}
                value={field.value ? [field.value] : []}
                onChange={(v) => field.onChange(v[0] ?? null)}
                placeholder="None"
              />
            )}
          />
        </div>
        <div className="form-section">
          <div className="form-section-title">Details</div>
          <div className="field-row">
            <SelectField
              label="Criticality"
              options={['low', 'medium', 'high', 'critical'].map((c) => ({ value: c, label: humanize(c) }))}
              {...form.register('criticality')}
            />
            {!asset && (
              <SelectField
                label="Initial status"
                options={ASSET_STATUSES.map((s) => ({ value: s, label: STATUS_LABEL[s] }))}
                {...form.register('status')}
              />
            )}
          </div>
          <div className="field-row">
            <TextField label="Manufacturer" {...form.register('manufacturer')} />
            <TextField label="Model" {...form.register('model')} />
          </div>
          <TextField label="Serial number" {...form.register('serial_number')} />
          <div className="field-row">
            <TextField label="Purchase date" type="date" {...form.register('purchase_date')} />
            <TextField
              label="Purchase cost (USD)"
              type="number"
              min={0}
              step="0.01"
              error={e.purchase_cost?.message}
              {...form.register('purchase_cost', { valueAsNumber: true })}
            />
          </div>
          <TextField label="Warranty ends" type="date" {...form.register('warranty_end')} />
        </div>
      </form>
    </SideSheet>
  );
}
