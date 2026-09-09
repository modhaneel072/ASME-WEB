import { useState } from 'react';
import { Link, useNavigate, useParams } from 'react-router-dom';
import { Controller, useForm } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import { AlertTriangle, Archive, ExternalLink, Pencil, Plus, Trash2, Users } from 'lucide-react';
import { api, ApiError } from '@/api/client';
import {
  keys,
  useInvalidate,
  useLookups,
  useProject,
  useProjectActivity,
  useProjectHealth,
} from '@/api/hooks';
import { useCurrentSession } from '@/auth/SessionProvider';
import { milestoneSchema, type MilestoneInput } from '@/contracts/schemas';
import type { Milestone, ProjectDetail, ProjectMember } from '@/contracts/types';
import { fmtDate, fmtMoney, humanize } from '@/lib/format';
import { userOptions } from '@/lib/options';
import { ActivityTimeline, auditItems } from '@/ui/ActivityTimeline';
import { AttachmentList, AttachmentUploader } from '@/ui/Attachments';
import {
  Avatar,
  Badge,
  Button,
  ConfirmDialog,
  Dialog,
  ErrorState,
  PageHeader,
  Picker,
  SelectField,
  Skeleton,
  Tabs,
  TextArea,
  TextField,
  useToast,
} from '@/ui';
import { KpiCard, BarList } from '@/ui/ReportCard';
import { WorkOrdersView } from '@/pages/work-orders/WorkOrdersView';
import { ProjectForm } from './ProjectForm';

type Tab = 'overview' | 'work-orders' | 'milestones' | 'members' | 'files' | 'activity';
const TABS: { key: Tab; label: string }[] = [
  { key: 'overview', label: 'Overview' },
  { key: 'work-orders', label: 'Work orders' },
  { key: 'milestones', label: 'Milestones' },
  { key: 'members', label: 'Members' },
  { key: 'files', label: 'Files' },
  { key: 'activity', label: 'Activity' },
];
const STATUS_TONE: Record<string, string> = {
  planning: 'info',
  active: 'success',
  on_hold: 'warning',
  completed: 'neutral',
  archived: 'neutral',
};

export function ProjectDetailPage() {
  const { id = '', tab: tabParam } = useParams();
  const tab = (TABS.some((t) => t.key === tabParam) ? tabParam : 'overview') as Tab;
  const navigate = useNavigate();
  const toast = useToast();
  const invalidate = useInvalidate();
  const { session } = useCurrentSession();
  const project = useProject(id);
  const [editing, setEditing] = useState(false);
  const [archiving, setArchiving] = useState(false);

  if (project.isPending)
    return (
      <div className="page-body" style={{ paddingTop: 24 }}>
        <Skeleton lines={6} />
      </div>
    );
  if (project.error || !project.data)
    return (
      <div className="page-body" style={{ paddingTop: 24 }}>
        <ErrorState
          error={project.error}
          onRetry={() => project.refetch()}
          title={
            project.error instanceof ApiError && project.error.status === 404
              ? 'Project not found'
              : undefined
          }
        />
      </div>
    );
  const p = project.data;
  const canEdit = p.permissions?.edit ?? false;

  return (
    <div className="page">
      <PageHeader
        breadcrumbs={[{ label: 'Projects', to: '/projects' }, { label: p.code }]}
        title={p.name}
        testId="project-detail"
        subtitle={
          <span className="row" style={{ gap: 8, flexWrap: 'wrap' }}>
            <Badge tone={STATUS_TONE[p.status] || 'neutral'}>{humanize(p.status)}</Badge>
            <Badge
              tone={p.risk_level === 'high' ? 'critical' : p.risk_level === 'medium' ? 'warning' : 'success'}
              icon={p.risk_level === 'high' ? <AlertTriangle /> : undefined}
            >
              {humanize(p.risk_level)} risk
            </Badge>
            {p.lead && (
              <span className="row" style={{ gap: 6 }}>
                <Avatar user={p.lead} size="sm" /> Lead: {p.lead.name}
              </span>
            )}
            {p.competition && <span>· {p.competition}</span>}
            {p.public_slug && (
              <a
                href={`/projects/${p.public_slug}`}
                target="_blank"
                rel="noreferrer"
                className="row"
                style={{ gap: 4 }}
              >
                Public page <ExternalLink size={12} />
              </a>
            )}
          </span>
        }
        actions={
          <>
            {p.permissions?.create_work_order && (
              <Button
                variant="primary"
                icon={<Plus />}
                onClick={() => navigate(`/projects/${p.id}/work-orders?new=1`)}
                data-testid="project-new-work-order"
              >
                New work order
              </Button>
            )}
            {canEdit && (
              <Button icon={<Pencil />} onClick={() => setEditing(true)} data-testid="project-edit">
                Edit
              </Button>
            )}
            {canEdit && !p.archived_at && (
              <Button variant="ghost" icon={<Archive />} onClick={() => setArchiving(true)}>
                Archive
              </Button>
            )}
          </>
        }
      />
      <div className="page-body">
        <Tabs
          label="Project sections"
          value={tab}
          tabs={TABS.map((t) => ({
            ...t,
            to: `/projects/${p.id}/${t.key === 'overview' ? '' : t.key}`,
            count:
              t.key === 'milestones'
                ? p.milestones.length
                : t.key === 'members'
                  ? p.members.length
                  : t.key === 'work-orders'
                    ? p.open_work_orders
                    : undefined,
          }))}
        />
        <div style={{ paddingTop: 20 }}>
          {tab === 'overview' && <Overview project={p} />}
          {tab === 'work-orders' && <WorkOrdersView projectId={p.id} embedded />}
          {tab === 'milestones' && <Milestones project={p} canEdit={canEdit} />}
          {tab === 'members' && <Members project={p} canEdit={canEdit} />}
          {tab === 'files' && (
            <div className="stack" style={{ maxWidth: 720 }}>
              <AttachmentList
                files={p.files || []}
                currentUser={session.user}
                canDelete={(f) => canEdit || f.uploaded_by?.id === session.user.id}
                onChanged={() => invalidate(keys.project(p.id))}
              />
              {(canEdit || session.permissions.includes('file.upload')) && (
                <AttachmentUploader
                  entityType="project"
                  entityId={p.id}
                  onUploaded={() => invalidate(keys.project(p.id))}
                />
              )}
            </div>
          )}
          {tab === 'activity' && <Activity projectId={p.id} />}
        </div>
      </div>
      {editing && (
        <ProjectForm
          project={p}
          onClose={() => setEditing(false)}
          onSaved={() => {
            setEditing(false);
          }}
        />
      )}
      <ConfirmDialog
        open={archiving}
        onClose={() => setArchiving(false)}
        title={`Archive ${p.name}?`}
        body="Archived projects are hidden from the active list and from selectors. Their work orders stay intact. This can be reversed by editing the project status."
        confirmLabel="Archive project"
        danger
        onConfirm={async () => {
          try {
            await api(`/projects/${p.id}/archive`, { method: 'POST' });
            invalidate(['projects'], keys.lookups);
            toast.success('Project archived');
            navigate('/projects');
          } catch (err) {
            toast.error('Could not archive', err instanceof ApiError ? err.message : undefined);
          }
          setArchiving(false);
        }}
      />
    </div>
  );
}

function Overview({ project: p }: { project: ProjectDetail }) {
  const health = useProjectHealth(p.id);
  const h = health.data;
  return (
    <div className="stack" style={{ gap: 20 }}>
      <div className="report-grid report-grid-kpi">
        <KpiCard
          label="Completion"
          value={`${p.completion_percent}%`}
          help="Weighted average of milestone status and done work orders."
        />
        <KpiCard label="Open work orders" value={p.open_work_orders} />
        <KpiCard
          label="Overdue"
          value={p.overdue_work_orders}
          tone={p.overdue_work_orders ? 'bad' : undefined}
        />
        <KpiCard label="Blocked" value={p.blocked_work_orders} />
        <KpiCard label="Hours logged" value={p.hours_logged} />
        <KpiCard
          label="Budget used"
          value={
            p.budget_amount
              ? `${fmtMoney(p.budget_used)} / ${fmtMoney(p.budget_amount)}`
              : fmtMoney(p.budget_used)
          }
          help="Sum of cost entries on this project's work orders."
        />
      </div>
      <div className="grid-2">
        <section className="card">
          <div className="card-head">
            <span className="card-title">About</span>
          </div>
          <div className="card-body stack">
            {p.description ? (
              <p className="description-block">{p.description}</p>
            ) : (
              <p className="text-muted">No description yet.</p>
            )}
            <dl className="kv">
              <dt>Season</dt>
              <dd>{[p.academic_year, p.competition].filter(Boolean).join(' · ') || '—'}</dd>
              <dt>Dates</dt>
              <dd>
                {fmtDate(p.start_date)} → {fmtDate(p.target_date)}
              </dd>
              <dt>Advisor</dt>
              <dd>{p.faculty_advisor?.name || '—'}</dd>
              <dt>Budget code</dt>
              <dd>{p.budget_code || '—'}</dd>
              <dt>Links</dt>
              <dd className="row" style={{ flexWrap: 'wrap' }}>
                {p.repository_url && (
                  <a href={p.repository_url} target="_blank" rel="noreferrer">
                    Repository
                  </a>
                )}
                {p.cad_url && (
                  <a href={p.cad_url} target="_blank" rel="noreferrer">
                    CAD
                  </a>
                )}
                {p.requirements_url && (
                  <a href={p.requirements_url} target="_blank" rel="noreferrer">
                    Requirements
                  </a>
                )}
                {!p.repository_url && !p.cad_url && !p.requirements_url && '—'}
              </dd>
              <dt>Visibility</dt>
              <dd>{humanize(p.visibility)}</dd>
            </dl>
          </div>
        </section>
        <section className="card">
          <div className="card-head">
            <span className="card-title">Teams</span>
            <Link to="/teams-users" className="text-label">
              Manage
            </Link>
          </div>
          <div className="card-body stack">
            {p.teams.length === 0 && (
              <span className="text-muted">No teams are attached to this project yet.</span>
            )}
            {p.teams.map((t) => (
              <div key={t.id} className="row-between">
                <span className="row">
                  <span
                    className="chip-dot"
                    style={{ background: t.color || 'var(--color-border-strong)' }}
                  />
                  {t.name}
                </span>
                <span className="text-caption text-muted">{t.member_count} members</span>
              </div>
            ))}
            {h && h.workload_by_team.length > 0 && (
              <>
                <div className="divider" />
                <div className="text-label">Open work by team</div>
                <BarList
                  rows={h.workload_by_team.map((w) => ({ label: w.team, value: w.open_work_orders }))}
                />
              </>
            )}
          </div>
        </section>
      </div>
      <section className="card">
        <div className="card-head">
          <span className="card-title">Milestones</span>
          <Link to={`/projects/${p.id}/milestones`} className="text-label">
            All milestones
          </Link>
        </div>
        <div className="card-body">
          {p.milestones.length === 0 ? (
            <span className="text-muted">No milestones yet.</span>
          ) : (
            <MilestoneList milestones={p.milestones} />
          )}
          {h && h.milestones.at_risk.length > 0 && (
            <div className="notice notice-warning" style={{ marginTop: 12 }}>
              <AlertTriangle />
              <div>
                {h.milestones.at_risk.length} milestone{h.milestones.at_risk.length === 1 ? ' is' : 's are'}{' '}
                overdue or due within two weeks: {h.milestones.at_risk.map((m) => m.name).join(', ')}.
              </div>
            </div>
          )}
        </div>
      </section>
    </div>
  );
}

const MS_TONE: Record<string, string> = {
  planned: 'neutral',
  in_progress: 'info',
  complete: 'success',
  missed: 'critical',
};

function MilestoneList({
  milestones,
  onEdit,
  onDelete,
}: {
  milestones: Milestone[];
  onEdit?: (m: Milestone) => void;
  onDelete?: (m: Milestone) => void;
}) {
  return (
    <div className="table-wrap">
      <table className="table">
        <thead>
          <tr>
            <th>Milestone</th>
            <th>Due</th>
            <th>Status</th>
            <th>Owner</th>
            <th className="num">Weight</th>
            {(onEdit || onDelete) && <th />}
          </tr>
        </thead>
        <tbody>
          {milestones.map((m) => (
            <tr key={m.id} data-testid="milestone-row">
              <td style={{ fontWeight: 600 }}>{m.name}</td>
              <td style={m.is_overdue ? { color: 'var(--color-danger)', fontWeight: 600 } : undefined}>
                {fmtDate(m.due_date)}
              </td>
              <td>
                <Badge tone={MS_TONE[m.status]}>{humanize(m.status)}</Badge>
              </td>
              <td>{m.owner?.name || '—'}</td>
              <td className="num">{m.weight}</td>
              {(onEdit || onDelete) && (
                <td>
                  <span className="inline-actions">
                    {onEdit && (
                      <Button
                        size="sm"
                        variant="ghost"
                        onClick={() => onEdit(m)}
                        aria-label={`Edit ${m.name}`}
                        icon={<Pencil />}
                      />
                    )}
                    {onDelete && (
                      <Button
                        size="sm"
                        variant="ghost"
                        onClick={() => onDelete(m)}
                        aria-label={`Delete ${m.name}`}
                        icon={<Trash2 />}
                      />
                    )}
                  </span>
                </td>
              )}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

function Milestones({ project: p, canEdit }: { project: ProjectDetail; canEdit: boolean }) {
  const [editing, setEditing] = useState<Milestone | 'new' | null>(null);
  const [deleting, setDeleting] = useState<Milestone | null>(null);
  const invalidate = useInvalidate();
  const toast = useToast();
  return (
    <div className="stack">
      <div className="row-between">
        <span className="text-secondary">
          Milestones drive the completion percentage; weight them by effort.
        </span>
        {canEdit && (
          <Button
            variant="primary"
            size="sm"
            icon={<Plus />}
            onClick={() => setEditing('new')}
            data-testid="add-milestone"
          >
            Add milestone
          </Button>
        )}
      </div>
      {p.milestones.length === 0 ? (
        <div className="card">
          <div className="card-body text-muted">No milestones yet.</div>
        </div>
      ) : (
        <MilestoneList
          milestones={p.milestones}
          onEdit={canEdit ? setEditing : undefined}
          onDelete={canEdit ? setDeleting : undefined}
        />
      )}
      {editing && (
        <MilestoneDialog
          project={p}
          milestone={editing === 'new' ? undefined : editing}
          onClose={() => setEditing(null)}
        />
      )}
      <ConfirmDialog
        open={!!deleting}
        onClose={() => setDeleting(null)}
        title="Delete milestone?"
        body={`"${deleting?.name}" will be removed from the plan.`}
        confirmLabel="Delete"
        danger
        onConfirm={async () => {
          if (!deleting) return;
          try {
            await api(`/projects/${p.id}/milestones/${deleting.id}`, { method: 'DELETE' });
            invalidate(keys.project(p.id), keys.projectHealth(p.id), ['projects']);
            toast.success('Milestone deleted');
          } catch (err) {
            toast.error('Could not delete', err instanceof ApiError ? err.message : undefined);
          }
          setDeleting(null);
        }}
      />
    </div>
  );
}

function MilestoneDialog({
  project,
  milestone,
  onClose,
}: {
  project: ProjectDetail;
  milestone?: Milestone;
  onClose: () => void;
}) {
  const lookups = useLookups();
  const invalidate = useInvalidate();
  const toast = useToast();
  const [banner, setBanner] = useState<string | null>(null);
  const form = useForm<MilestoneInput>({
    resolver: zodResolver(milestoneSchema),
    defaultValues: {
      name: milestone?.name || '',
      description: milestone?.description || '',
      due_date: milestone?.due_date || '',
      status: milestone?.status || 'planned',
      owner_user_id: milestone?.owner?.id ?? null,
      weight: milestone?.weight ?? 1,
    },
  });
  const submit = form.handleSubmit(async (values) => {
    setBanner(null);
    const data = milestoneSchema.parse(values);
    try {
      if (milestone)
        await api(`/projects/${project.id}/milestones/${milestone.id}`, { method: 'PATCH', body: data });
      else await api(`/projects/${project.id}/milestones`, { method: 'POST', body: data });
      invalidate(keys.project(project.id), keys.projectHealth(project.id), ['projects']);
      toast.success(milestone ? 'Milestone updated' : 'Milestone added');
      onClose();
    } catch (err) {
      if (err instanceof ApiError) {
        const fields = err.fieldErrors();
        Object.entries(fields).forEach(([f, m]) => form.setError(f as keyof MilestoneInput, { message: m }));
        if (!Object.keys(fields).length) setBanner(err.message);
      }
    }
  });
  const e = form.formState.errors;
  return (
    <Dialog
      open
      onClose={onClose}
      title={milestone ? 'Edit milestone' : 'Add milestone'}
      footer={
        <>
          <Button variant="ghost" onClick={onClose}>
            Cancel
          </Button>
          <Button
            variant="primary"
            onClick={submit}
            loading={form.formState.isSubmitting}
            data-testid="milestone-submit"
          >
            {milestone ? 'Save' : 'Add milestone'}
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
          label="Name"
          required
          error={e.name?.message}
          {...form.register('name')}
          data-autofocus
          data-testid="milestone-name"
        />
        <TextArea label="Description" rows={2} {...form.register('description')} />
        <div className="field-row">
          <TextField label="Due date" type="date" {...form.register('due_date')} />
          <SelectField
            label="Status"
            options={['planned', 'in_progress', 'complete', 'missed'].map((s) => ({
              value: s,
              label: humanize(s),
            }))}
            {...form.register('status')}
          />
        </div>
        <div className="field-row">
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
                placeholder="Unowned"
              />
            )}
          />
          <TextField
            label="Weight"
            type="number"
            min={1}
            max={100}
            error={e.weight?.message}
            {...form.register('weight', { valueAsNumber: true })}
            hint="Relative effort; used for completion %."
          />
        </div>
      </form>
    </Dialog>
  );
}

function Members({ project: p, canEdit }: { project: ProjectDetail; canEdit: boolean }) {
  const [managing, setManaging] = useState(false);
  return (
    <div className="stack">
      <div className="row-between">
        <span className="text-secondary">
          Members see this project even when it is private; the lead can manage it.
        </span>
        {canEdit && (
          <Button
            variant="primary"
            size="sm"
            icon={<Users />}
            onClick={() => setManaging(true)}
            data-testid="manage-members"
          >
            Manage members
          </Button>
        )}
      </div>
      <div className="table-wrap">
        <table className="table">
          <thead>
            <tr>
              <th>Member</th>
              <th>Project role</th>
              <th>Team</th>
              <th>Chapter role</th>
            </tr>
          </thead>
          <tbody>
            {p.members.length === 0 && (
              <tr className="table-empty">
                <td colSpan={4} className="text-muted">
                  No members yet.
                </td>
              </tr>
            )}
            {p.members.map((m) => (
              <tr key={m.id} data-testid="member-row">
                <td>
                  <span className="row">
                    <Avatar user={m} size="sm" /> {m.name}
                  </span>
                </td>
                <td>
                  <Badge
                    tone={
                      m.project_role === 'lead'
                        ? 'info'
                        : m.project_role === 'advisor'
                          ? 'warning'
                          : 'neutral'
                    }
                  >
                    {humanize(m.project_role)}
                  </Badge>
                </td>
                <td>{m.team?.name || '—'}</td>
                <td>{m.role_name || '—'}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
      {managing && <MembersDialog project={p} onClose={() => setManaging(false)} />}
    </div>
  );
}

function MembersDialog({ project, onClose }: { project: ProjectDetail; onClose: () => void }) {
  const lookups = useLookups();
  const invalidate = useInvalidate();
  const toast = useToast();
  const [rows, setRows] = useState<
    { user_id: number; project_role: ProjectMember['project_role']; team_id: string | null }[]
  >(project.members.map((m) => ({ user_id: m.id, project_role: m.project_role, team_id: m.team_id })));
  const [busy, setBusy] = useState(false);
  const selectedIds = rows.map((r) => r.user_id);
  const users = lookups.data?.users || [];
  const teams = (lookups.data?.teams || []).filter((t) => !t.project_id || t.project_id === project.id);
  return (
    <Dialog
      open
      onClose={onClose}
      size="lg"
      title="Project members"
      description="Pick who is on the project, then set their project role and team."
      footer={
        <>
          <Button variant="ghost" onClick={onClose}>
            Cancel
          </Button>
          <Button
            variant="primary"
            loading={busy}
            data-testid="members-submit"
            onClick={async () => {
              setBusy(true);
              try {
                await api(`/projects/${project.id}/members`, { method: 'PUT', body: { members: rows } });
                invalidate(keys.project(project.id), ['projects']);
                toast.success('Members updated');
                onClose();
              } catch (err) {
                toast.error('Could not save members', err instanceof ApiError ? err.message : undefined);
              } finally {
                setBusy(false);
              }
            }}
          >
            Save members
          </Button>
        </>
      }
    >
      <div className="stack" style={{ gap: 16 }}>
        <Picker
          label="Members"
          options={userOptions(lookups.data)}
          value={selectedIds}
          onChange={(ids) =>
            setRows((prev) =>
              ids.map(
                (id) =>
                  prev.find((r) => r.user_id === id) || {
                    user_id: id,
                    project_role: 'member',
                    team_id: null,
                  },
              ),
            )
          }
          placeholder="Add people"
          id="project-members"
        />
        {rows.length > 0 && (
          <div className="table-wrap">
            <table className="table">
              <thead>
                <tr>
                  <th>Member</th>
                  <th>Project role</th>
                  <th>Team</th>
                </tr>
              </thead>
              <tbody>
                {rows.map((r, i) => {
                  const user = users.find((u) => u.id === r.user_id);
                  return (
                    <tr key={r.user_id}>
                      <td>{user?.name || r.user_id}</td>
                      <td>
                        <select
                          className="field-control"
                          style={{ height: 32 }}
                          value={r.project_role}
                          aria-label={`Project role for ${user?.name}`}
                          onChange={(e) =>
                            setRows((prev) =>
                              prev.map((x, j) =>
                                j === i
                                  ? { ...x, project_role: e.target.value as ProjectMember['project_role'] }
                                  : x,
                              ),
                            )
                          }
                        >
                          {['lead', 'member', 'advisor', 'observer'].map((role) => (
                            <option key={role} value={role}>
                              {humanize(role)}
                            </option>
                          ))}
                        </select>
                      </td>
                      <td>
                        <select
                          className="field-control"
                          style={{ height: 32 }}
                          value={r.team_id || ''}
                          aria-label={`Team for ${user?.name}`}
                          onChange={(e) =>
                            setRows((prev) =>
                              prev.map((x, j) => (j === i ? { ...x, team_id: e.target.value || null } : x)),
                            )
                          }
                        >
                          <option value="">No team</option>
                          {teams.map((t) => (
                            <option key={t.id} value={t.id}>
                              {t.name}
                            </option>
                          ))}
                        </select>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </Dialog>
  );
}

function Activity({ projectId }: { projectId: string }) {
  const activity = useProjectActivity(projectId);
  if (activity.isPending) return <Skeleton lines={5} />;
  if (activity.error) return <ErrorState error={activity.error} onRetry={() => activity.refetch()} />;
  return (
    <div className="card">
      <div className="card-body">
        <ActivityTimeline
          items={auditItems(activity.data || [])}
          emptyText="No activity recorded for this project yet."
        />
      </div>
    </div>
  );
}
