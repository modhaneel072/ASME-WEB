import { useMemo, useState } from 'react';
import { useNavigate, useParams } from 'react-router-dom';
import { Controller, useForm } from 'react-hook-form';
import { zodResolver } from '@hookform/resolvers/zod';
import { Archive, Plus, UserPlus, Users } from 'lucide-react';
import { api, ApiError } from '@/api/client';
import { keys, useInvalidate, useLookups, useRoles, useTeams, useUsers } from '@/api/hooks';
import { useCurrentSession } from '@/auth/SessionProvider';
import { inviteSchema, teamSchema, type InviteInput, type TeamInput } from '@/contracts/schemas';
import type { Member, Team } from '@/contracts/types';
import { fmtDate, humanize } from '@/lib/format';
import { projectOptions, userOptions } from '@/lib/options';
import {
  Avatar,
  AvatarStack,
  Badge,
  Button,
  Checkbox,
  ConfirmDialog,
  EmptyState,
  ErrorState,
  PageHeader,
  Picker,
  SearchField,
  SelectField,
  SideSheet,
  Skeleton,
  Tabs,
  TextArea,
  TextField,
  useToast,
} from '@/ui';

export function TeamsUsersPage() {
  const { tab = 'teams' } = useParams();
  const navigate = useNavigate();
  const { can } = useCurrentSession();
  return (
    <div className="page">
      <PageHeader
        title="Teams & Users"
        subtitle="Teams route work; roles decide what people can do."
        testId="teams-users-page"
      />
      <div className="page-body">
        <Tabs
          label="Teams or users"
          value={tab}
          tabs={[
            { key: 'teams', label: 'Teams', to: '/teams-users' },
            { key: 'users', label: 'Users', to: '/teams-users/users' },
          ]}
          onChange={(k) => navigate(k === 'teams' ? '/teams-users' : '/teams-users/users')}
        />
        <div style={{ paddingTop: 20 }}>
          {tab === 'users' ? (
            <UsersTab canManage={can('user.manage')} />
          ) : (
            <TeamsTab canManage={can('team.manage')} />
          )}
        </div>
      </div>
    </div>
  );
}

// ---------------------------------------------------------------------------- teams

function TeamsTab({ canManage }: { canManage: boolean }) {
  const teams = useTeams();
  const [editing, setEditing] = useState<Team | 'new' | null>(null);
  const [archiving, setArchiving] = useState<Team | null>(null);
  const [q, setQ] = useState('');
  const invalidate = useInvalidate();
  const toast = useToast();
  const rows = useMemo(
    () => (teams.data || []).filter((t) => !q || t.name.toLowerCase().includes(q.toLowerCase())),
    [teams.data, q],
  );
  if (teams.isPending) return <Skeleton lines={5} />;
  if (teams.error) return <ErrorState error={teams.error} onRetry={() => teams.refetch()} />;
  return (
    <div className="stack">
      <div className="row-between">
        <SearchField value={q} onChange={setQ} placeholder="Search teams" />
        {canManage && (
          <Button variant="primary" icon={<Plus />} onClick={() => setEditing('new')} data-testid="new-team">
            New team
          </Button>
        )}
      </div>
      {rows.length === 0 ? (
        <div className="card">
          <EmptyState
            icon={<Users />}
            title={q ? 'No teams match' : 'No teams yet'}
            body="Teams group members so work can be assigned to a whole subteam."
            actions={
              canManage && !q ? (
                <Button variant="primary" icon={<Plus />} onClick={() => setEditing('new')}>
                  New team
                </Button>
              ) : undefined
            }
          />
        </div>
      ) : (
        <div className="table-wrap">
          <table className="table" data-testid="teams-table">
            <thead>
              <tr>
                <th>Team</th>
                <th>Project</th>
                <th>Parent</th>
                <th>Leads</th>
                <th>Members</th>
                {canManage && <th />}
              </tr>
            </thead>
            <tbody>
              {rows.map((t) => (
                <tr
                  key={t.id}
                  className={canManage ? 'clickable' : undefined}
                  onClick={() => canManage && setEditing(t)}
                  data-testid="team-row"
                >
                  <td>
                    <span className="row">
                      <span
                        className="chip-dot"
                        style={{ background: t.color || 'var(--color-border-strong)' }}
                      />
                      <span style={{ fontWeight: 600 }}>{t.name}</span>
                    </span>
                    {t.description && <div className="text-caption text-muted">{t.description}</div>}
                  </td>
                  <td>{t.project?.name || <span className="text-muted">Chapter-wide</span>}</td>
                  <td>{t.parent?.name || '—'}</td>
                  <td>
                    {t.members
                      .filter((m) => m.is_lead)
                      .map((m) => m.name)
                      .join(', ') || '—'}
                  </td>
                  <td>
                    <AvatarStack users={t.members} max={4} />
                  </td>
                  {canManage && (
                    <td onClick={(e) => e.stopPropagation()}>
                      <Button
                        size="sm"
                        variant="ghost"
                        icon={<Archive />}
                        onClick={() => setArchiving(t)}
                        aria-label={`Archive ${t.name}`}
                      />
                    </td>
                  )}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
      {editing && (
        <TeamForm team={editing === 'new' ? undefined : editing} onClose={() => setEditing(null)} />
      )}
      <ConfirmDialog
        open={!!archiving}
        onClose={() => setArchiving(null)}
        title={`Archive ${archiving?.name}?`}
        body="The team disappears from selectors. Existing assignments are kept."
        confirmLabel="Archive"
        danger
        onConfirm={async () => {
          if (!archiving) return;
          try {
            await api(`/teams/${archiving.id}`, { method: 'DELETE' });
            invalidate(keys.teams, keys.lookups, keys.setup);
            toast.success('Team archived');
          } catch (err) {
            toast.error('Could not archive', err instanceof ApiError ? err.message : undefined);
          }
          setArchiving(null);
        }}
      />
    </div>
  );
}

function TeamForm({ team, onClose }: { team?: Team; onClose: () => void }) {
  const lookups = useLookups();
  const teams = useTeams();
  const invalidate = useInvalidate();
  const toast = useToast();
  const [banner, setBanner] = useState<string | null>(null);
  const [memberIds, setMemberIds] = useState<number[]>(team?.members.map((m) => m.id) || []);
  const [leadIds, setLeadIds] = useState<number[]>(
    team?.members.filter((m) => m.is_lead).map((m) => m.id) || [],
  );
  const form = useForm<TeamInput>({
    resolver: zodResolver(teamSchema),
    defaultValues: {
      name: team?.name || '',
      description: team?.description || '',
      parent_team_id: team?.parent_team_id || '',
      project_id: team?.project_id || '',
      color: team?.color || '#0878d1',
      escalation_note: team?.escalation_note || '',
    },
  });
  const users = lookups.data?.users || [];
  const submit = form.handleSubmit(async (values) => {
    setBanner(null);
    const data = teamSchema.parse(values);
    const members = memberIds.map((id) => ({ user_id: id, is_lead: leadIds.includes(id) }));
    try {
      if (team) {
        const body: Record<string, unknown> = {};
        const dirty = form.formState.dirtyFields as Record<string, unknown>;
        for (const key of Object.keys(data) as (keyof TeamInput)[])
          if (dirty[key]) body[key] = data[key] ?? null;
        if (Object.keys(body).length) await api(`/teams/${team.id}`, { method: 'PATCH', body });
        await api(`/teams/${team.id}/members`, { method: 'PUT', body: { members } });
        toast.success('Team updated');
      } else {
        await api('/teams', { method: 'POST', body: { ...data, members } });
        toast.success('Team created');
      }
      invalidate(keys.teams, keys.lookups, keys.setup, ['projects'], ['project']);
      onClose();
    } catch (err) {
      if (err instanceof ApiError) {
        const fields = err.fieldErrors();
        Object.entries(fields).forEach(([f, m]) => form.setError(f as keyof TeamInput, { message: m }));
        if (!Object.keys(fields).length) setBanner(err.message);
      }
    }
  });
  const e = form.formState.errors;
  return (
    <SideSheet
      open
      onClose={onClose}
      title={team ? `Edit ${team.name}` : 'New team'}
      testId="team-form"
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
            data-testid="team-submit"
          >
            {team ? 'Save changes' : 'Create team'}
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
        <TextField
          label="Name"
          required
          error={e.name?.message}
          {...form.register('name')}
          data-autofocus
          data-testid="team-name"
        />
        <TextArea label="Description" rows={2} {...form.register('description')} />
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
                placeholder="Chapter-wide team"
              />
            )}
          />
          <Controller
            control={form.control}
            name="parent_team_id"
            render={({ field }) => (
              <Picker
                label="Parent team"
                multiple={false}
                options={(teams.data || [])
                  .filter((t) => t.id !== team?.id)
                  .map((t) => ({ value: t.id, label: t.name, color: t.color }))}
                value={field.value ? [field.value] : []}
                onChange={(v) => field.onChange(v[0] || '')}
                placeholder="None"
              />
            )}
          />
        </div>
        <div className="field-row">
          <TextField
            label="Colour"
            type="color"
            {...form.register('color')}
            style={{ padding: 4, width: 80 }}
          />
          <TextField
            label="Escalation note"
            placeholder="Who to contact when this team is stuck"
            {...form.register('escalation_note')}
          />
        </div>
        <Picker
          label="Members"
          options={userOptions(lookups.data)}
          value={memberIds}
          onChange={(ids) => {
            setMemberIds(ids);
            setLeadIds((prev) => prev.filter((l) => ids.includes(l)));
          }}
          placeholder="Add members"
          id="team-members"
        />
        {memberIds.length > 0 && (
          <div className="field">
            <span className="field-label">Team leads</span>
            <div className="stack" style={{ gap: 6 }}>
              {memberIds.map((id) => {
                const u = users.find((x) => x.id === id);
                return (
                  <Checkbox
                    key={id}
                    label={
                      <span className="row">
                        <Avatar user={u || { id, name: String(id), initials: '?' }} size="sm" />{' '}
                        {u?.name || id}
                      </span>
                    }
                    checked={leadIds.includes(id)}
                    onChange={(ev) =>
                      setLeadIds((prev) => (ev.target.checked ? [...prev, id] : prev.filter((x) => x !== id)))
                    }
                  />
                );
              })}
            </div>
            <div className="field-hint">Leads can assign and edit work for their team.</div>
          </div>
        )}
      </form>
    </SideSheet>
  );
}

// ---------------------------------------------------------------------------- users

function UsersTab({ canManage }: { canManage: boolean }) {
  const [includeInactive, setIncludeInactive] = useState(false);
  const users = useUsers(includeInactive);
  const roles = useRoles();
  const [q, setQ] = useState('');
  const [inviting, setInviting] = useState(false);
  const { session } = useCurrentSession();
  const rows = useMemo(
    () =>
      (users.data || []).filter(
        (u) =>
          !q ||
          u.name.toLowerCase().includes(q.toLowerCase()) ||
          u.email.toLowerCase().includes(q.toLowerCase()),
      ),
    [users.data, q],
  );
  if (users.isPending) return <Skeleton lines={6} />;
  if (users.error) return <ErrorState error={users.error} onRetry={() => users.refetch()} />;
  return (
    <div className="stack">
      <div className="row-between" style={{ flexWrap: 'wrap' }}>
        <div className="row">
          <SearchField value={q} onChange={setQ} placeholder="Search name or email" />
          {canManage && (
            <Checkbox
              label="Include inactive"
              checked={includeInactive}
              onChange={(e) => setIncludeInactive(e.target.checked)}
            />
          )}
        </div>
        {canManage && (
          <Button
            variant="primary"
            icon={<UserPlus />}
            onClick={() => setInviting(true)}
            data-testid="invite-user"
          >
            Invite member
          </Button>
        )}
      </div>
      <div className="table-wrap">
        <table className="table" data-testid="users-table">
          <thead>
            <tr>
              <th>Member</th>
              <th>Role</th>
              <th>Status</th>
              <th>Title</th>
              <th>Last sign-in</th>
            </tr>
          </thead>
          <tbody>
            {rows.map((u) => (
              <UserRow
                key={u.id}
                user={u}
                roles={roles.data || []}
                canManage={canManage && u.id !== session.user.id}
              />
            ))}
          </tbody>
        </table>
      </div>
      {inviting && <InviteForm onClose={() => setInviting(false)} />}
    </div>
  );
}

function UserRow({
  user,
  roles,
  canManage,
}: {
  user: Member;
  roles: { key: string | null; name: string }[];
  canManage: boolean;
}) {
  const invalidate = useInvalidate();
  const toast = useToast();
  const [busy, setBusy] = useState(false);
  const patch = async (body: Record<string, unknown>) => {
    setBusy(true);
    try {
      await api(`/users/${user.id}`, { method: 'PATCH', body });
      invalidate(['users'], keys.lookups, keys.setup);
      toast.success(`${user.name} updated`);
    } catch (err) {
      toast.error('Could not update member', err instanceof ApiError ? err.message : undefined);
    } finally {
      setBusy(false);
    }
  };
  return (
    <tr data-testid="user-row">
      <td>
        <span className="row">
          <Avatar user={user} size="sm" />
          <span>
            <span style={{ fontWeight: 600, display: 'block' }}>{user.name}</span>
            <span className="text-caption text-muted">{user.email}</span>
          </span>
        </span>
      </td>
      <td>
        {canManage ? (
          <select
            className="field-control"
            style={{ height: 32, minWidth: 180 }}
            value={user.role_key || ''}
            disabled={busy}
            aria-label={`Role for ${user.name}`}
            onChange={(e) => patch({ role_key: e.target.value })}
            data-testid="role-select"
          >
            {roles
              .filter((r) => r.key)
              .map((r) => (
                <option key={r.key!} value={r.key!}>
                  {r.name}
                </option>
              ))}
          </select>
        ) : (
          user.role_name || '—'
        )}
      </td>
      <td>
        {canManage ? (
          <select
            className="field-control"
            style={{ height: 32 }}
            value={user.member_status}
            disabled={busy}
            aria-label={`Status for ${user.name}`}
            onChange={(e) => patch({ member_status: e.target.value })}
          >
            {['active', 'invited', 'suspended'].map((s) => (
              <option key={s} value={s}>
                {humanize(s)}
              </option>
            ))}
          </select>
        ) : (
          <Badge
            tone={
              user.member_status === 'active'
                ? 'success'
                : user.member_status === 'invited'
                  ? 'info'
                  : 'warning'
            }
          >
            {humanize(user.member_status)}
          </Badge>
        )}
        {!user.is_active && <span className="text-caption text-muted"> · account disabled</span>}
      </td>
      <td>{user.title || '—'}</td>
      <td>{user.last_login_at ? fmtDate(user.last_login_at, true) : 'Never'}</td>
    </tr>
  );
}

function InviteForm({ onClose }: { onClose: () => void }) {
  const roles = useRoles();
  const invalidate = useInvalidate();
  const toast = useToast();
  const [banner, setBanner] = useState<string | null>(null);
  const form = useForm<InviteInput>({
    resolver: zodResolver(inviteSchema),
    defaultValues: { name: '', email: '', role_key: 'full_member', password: '', title: '' },
  });
  const submit = form.handleSubmit(async (values) => {
    setBanner(null);
    const data = inviteSchema.parse(values);
    try {
      await api('/users', { method: 'POST', body: data });
      invalidate(['users'], keys.lookups, keys.setup);
      toast.success(
        'Member added',
        data.password
          ? 'Share the temporary password with them directly.'
          : 'They can sign in once a password is set.',
      );
      onClose();
    } catch (err) {
      if (err instanceof ApiError) {
        const fields = err.fieldErrors();
        Object.entries(fields).forEach(([f, m]) => form.setError(f as keyof InviteInput, { message: m }));
        if (!Object.keys(fields).length) setBanner(err.message);
      }
    }
  });
  const e = form.formState.errors;
  return (
    <SideSheet
      open
      onClose={onClose}
      title="Invite member"
      subtitle="Creates the account (or attaches an existing public-site account) and sets the chapter role."
      testId="invite-form"
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
            data-testid="invite-submit"
          >
            Add member
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
        <TextField
          label="Full name"
          required
          error={e.name?.message}
          {...form.register('name')}
          data-autofocus
          data-testid="invite-name"
        />
        <TextField
          label="Email"
          type="email"
          required
          error={e.email?.message}
          {...form.register('email')}
          data-testid="invite-email"
        />
        <SelectField
          label="Role"
          required
          options={(roles.data || []).filter((r) => r.key).map((r) => ({ value: r.key!, label: r.name }))}
          error={e.role_key?.message}
          {...form.register('role_key')}
        />
        <TextField label="Title" placeholder="e.g. Treasurer, Wheels lead" {...form.register('title')} />
        <TextField
          label="Temporary password"
          type="password"
          autoComplete="new-password"
          hint="Optional. Leave blank for existing accounts; the member keeps their current password."
          error={e.password?.message}
          {...form.register('password')}
        />
      </form>
    </SideSheet>
  );
}
