import { useState } from 'react';
import { useNavigate, useParams } from 'react-router-dom';
import { useForm } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import { z } from 'zod';
import { Save } from 'lucide-react';
import { api, ApiError } from '@/api/client';
import { keys, useAudit, useInvalidate, useLookups, useRoles } from '@/api/hooks';
import { useCurrentSession } from '@/auth/SessionProvider';
import { fmtDate } from '@/lib/format';
import { ActivityTimeline, auditItems } from '@/ui/ActivityTimeline';
import {
  Button,
  ErrorState,
  FilterChip,
  PageHeader,
  SelectField,
  Skeleton,
  Tabs,
  TextField,
  useToast,
} from '@/ui';

type Tab = 'chapter' | 'roles' | 'audit';

export function SettingsPage() {
  const { tab: tabParam } = useParams();
  const navigate = useNavigate();
  const { can } = useCurrentSession();
  const tabs = [
    { key: 'chapter' as Tab, label: 'Chapter', show: can('org.read') || can('org.manage') },
    { key: 'roles' as Tab, label: 'Roles & permissions', show: true },
    { key: 'audit' as Tab, label: 'Audit log', show: can('audit.read') },
  ].filter((t) => t.show);
  const tab = (tabs.some((t) => t.key === tabParam) ? tabParam : tabs[0]?.key) as Tab;
  return (
    <div className="page">
      <PageHeader
        title="Settings"
        subtitle="Chapter profile, roles and the audit trail."
        testId="settings-page"
      />
      <div className="page-body">
        <Tabs
          label="Settings sections"
          value={tab}
          tabs={tabs.map((t) => ({ key: t.key, label: t.label, to: `/settings/${t.key}` }))}
          onChange={(k) => navigate(`/settings/${k}`)}
        />
        <div style={{ paddingTop: 20, maxWidth: tab === 'chapter' ? 640 : undefined }}>
          {tab === 'chapter' && <ChapterProfile canManage={can('org.manage')} />}
          {tab === 'roles' && <RolesMatrix />}
          {tab === 'audit' && <AuditLog />}
        </div>
      </div>
    </div>
  );
}

const profileSchema = z.object({
  name: z.string().trim().min(2, 'Name must be at least 2 characters').max(200),
  timezone: z.string().trim().min(2).max(60),
  academic_year_start_month: z.number().int().min(1).max(12),
  logo_url: z.string().trim().max(500).optional().or(z.literal('')),
});
type ProfileInput = z.infer<typeof profileSchema>;

const MONTHS = [
  'January',
  'February',
  'March',
  'April',
  'May',
  'June',
  'July',
  'August',
  'September',
  'October',
  'November',
  'December',
];

function ChapterProfile({ canManage }: { canManage: boolean }) {
  const { session, refresh } = useCurrentSession();
  const invalidate = useInvalidate();
  const toast = useToast();
  const [banner, setBanner] = useState<string | null>(null);
  const org = session.organization;
  const form = useForm<ProfileInput>({
    resolver: zodResolver(profileSchema),
    defaultValues: {
      name: org.name,
      timezone: org.timezone,
      academic_year_start_month: org.academic_year_start_month,
      logo_url: org.logo_url || '',
    },
  });
  const submit = form.handleSubmit(async (values) => {
    setBanner(null);
    try {
      await api('/org', { method: 'PATCH', body: { ...values, logo_url: values.logo_url || null } });
      await refresh();
      invalidate(keys.setup);
      toast.success('Chapter profile saved');
      form.reset(values);
    } catch (err) {
      if (err instanceof ApiError) {
        const fields = err.fieldErrors();
        Object.entries(fields).forEach(([f, m]) => form.setError(f as keyof ProfileInput, { message: m }));
        if (!Object.keys(fields).length) setBanner(err.message);
      }
    }
  });
  const e = form.formState.errors;
  return (
    <form className="card" onSubmit={submit} noValidate data-testid="chapter-form">
      <div className="card-head">
        <span className="card-title">Chapter profile</span>
        {canManage && (
          <Button
            type="submit"
            variant="primary"
            icon={<Save />}
            loading={form.formState.isSubmitting}
            data-testid="chapter-save"
          >
            Save
          </Button>
        )}
      </div>
      <div className="card-body stack" style={{ gap: 16 }}>
        {!canManage && <div className="notice">Only chapter admins can change these settings.</div>}
        {banner && (
          <div className="form-banner form-banner-error" role="alert">
            {banner}
          </div>
        )}
        <TextField
          label="Chapter name"
          required
          disabled={!canManage}
          error={e.name?.message}
          {...form.register('name')}
        />
        <div className="field-row">
          <TextField
            label="Timezone"
            required
            disabled={!canManage}
            hint="IANA name, e.g. America/Chicago"
            error={e.timezone?.message}
            {...form.register('timezone')}
          />
          <SelectField
            label="Academic year starts"
            disabled={!canManage}
            options={MONTHS.map((m, i) => ({ value: String(i + 1), label: m }))}
            {...form.register('academic_year_start_month', { valueAsNumber: true })}
          />
        </div>
        <TextField
          label="Logo URL"
          disabled={!canManage}
          placeholder="https://…"
          error={e.logo_url?.message}
          {...form.register('logo_url')}
        />
        <div className="text-caption text-muted">
          Saving confirms the profile for the Setup Center checklist. Slug: {org.slug}
        </div>
      </div>
    </form>
  );
}

function RolesMatrix() {
  const roles = useRoles();
  if (roles.isPending) return <Skeleton lines={6} />;
  if (roles.error) return <ErrorState error={roles.error} onRetry={() => roles.refetch()} />;
  const list = roles.data || [];
  const permissions = Array.from(new Set(list.flatMap((r) => r.permissions.map(([key]) => key)))).sort();
  return (
    <div className="stack">
      <p className="text-secondary" style={{ maxWidth: 720 }}>
        Roles are enforced on the server for every request. A scope narrows where a permission applies:{' '}
        <strong>chapter</strong> (everything), <strong>project</strong> (projects you lead or belong to),{' '}
        <strong>team</strong> (your teams), <strong>assigned</strong> (work assigned to you),{' '}
        <strong>own</strong> (things you created).
      </p>
      <div className="table-wrap">
        <table className="table role-matrix" data-testid="roles-table">
          <thead>
            <tr>
              <th>Permission</th>
              {list.map((r) => (
                <th key={r.id}>{r.name}</th>
              ))}
            </tr>
          </thead>
          <tbody>
            {permissions.map((perm) => (
              <tr key={perm}>
                <td className="mono">{perm}</td>
                {list.map((r) => {
                  const scope = r.permissions.find(([key]) => key === perm)?.[1];
                  return (
                    <td key={r.id} className={scope ? 'yes' : 'scope'}>
                      {scope ? (
                        scope === 'chapter' ? (
                          '✓'
                        ) : (
                          <span title={`scope: ${scope}`}>
                            ✓ <span className="scope">{scope}</span>
                          </span>
                        )
                      ) : (
                        ''
                      )}
                    </td>
                  );
                })}
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </div>
  );
}

const ENTITY_TYPES = [
  'work_order',
  'project',
  'milestone',
  'asset',
  'team',
  'location',
  'category',
  'user',
  'membership',
  'organization',
  'comment',
  'attachment',
  'saved_filter',
];

function AuditLog() {
  const [entityType, setEntityType] = useState<string[]>([]);
  const [actor, setActor] = useState<string[]>([]);
  const [offset, setOffset] = useState(0);
  const lookups = useLookups();
  const audit = useAudit({
    entity_type: entityType[0],
    actor_id: actor[0],
    offset: offset ? String(offset) : undefined,
    'page[limit]': '50',
  });
  return (
    <div className="stack">
      <div className="filter-bar">
        <FilterChip
          label="Entity"
          options={ENTITY_TYPES.map((t) => ({ value: t, label: t.replace(/_/g, ' ') }))}
          value={entityType}
          onChange={(v) => {
            setEntityType(v);
            setOffset(0);
          }}
          single
        />
        <FilterChip
          label="Actor"
          options={(lookups.data?.users || []).map((u) => ({ value: String(u.id), label: u.name }))}
          value={actor}
          onChange={(v) => {
            setActor(v);
            setOffset(0);
          }}
          single
        />
      </div>
      {audit.isPending ? (
        <Skeleton lines={6} />
      ) : audit.error ? (
        <ErrorState error={audit.error} onRetry={() => audit.refetch()} />
      ) : (
        <div className="card">
          <div className="card-body">
            <ActivityTimeline
              items={auditItems(audit.data?.items || []).map((i, idx) => ({
                ...i,
                meta: `${i.meta} · ${audit.data!.items[idx].entity_type} ${audit.data!.items[idx].entity_id.slice(0, 8)}`,
              }))}
              emptyText="No events match."
            />
          </div>
          <div
            className="row-between"
            style={{ padding: '8px 16px', borderTop: '1px solid var(--color-border)' }}
          >
            <Button
              size="sm"
              variant="ghost"
              disabled={offset === 0}
              onClick={() => setOffset(Math.max(0, offset - 50))}
            >
              Newer
            </Button>
            <span className="text-caption text-muted">
              {audit.data?.items.length ? `${offset + 1}–${offset + audit.data.items.length}` : ''}{' '}
              {audit.data?.items[0] ? `· latest ${fmtDate(audit.data.items[0].occurred_at, true)}` : ''}
            </span>
            <Button
              size="sm"
              variant="ghost"
              disabled={audit.data?.next_offset == null}
              onClick={() => setOffset(audit.data?.next_offset || 0)}
            >
              Older
            </Button>
          </div>
        </div>
      )}
    </div>
  );
}
