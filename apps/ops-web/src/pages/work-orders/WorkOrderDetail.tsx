import { useState, type ReactNode } from 'react';
import { Link } from 'react-router-dom';
import { useQuery, useQueryClient } from '@tanstack/react-query';
import {
  ArrowLeft,
  Ban,
  CheckCircle2,
  Copy,
  DollarSign,
  Eye,
  GitBranch,
  Link2,
  MoreHorizontal,
  Pause,
  Pencil,
  Play,
  Plus,
  Repeat,
  ShieldAlert,
  Timer,
  UserPlus,
  X,
} from 'lucide-react';
import { api, ApiError } from '@/api/client';
import { keys, useInvalidate, useWorkOrder } from '@/api/hooks';
import type {
  AuditEvent,
  Comment,
  StatusHistoryEntry,
  UserRef,
  WorkOrderDetail as WorkOrderDetailType,
  WorkOrderListItem,
} from '@/contracts/types';
import { fmtDate, fmtDue, fmtMinutes, fmtMoney, humanize, STATUS_LABEL } from '@/lib/format';
import { ActivityTimeline, auditItems, statusHistoryItems } from '@/ui/ActivityTimeline';
import { AttachmentList, AttachmentUploader } from '@/ui/Attachments';
import { CommentThread } from '@/ui/Comments';
import {
  Avatar,
  AvatarStack,
  BlockedBadge,
  Button,
  CategoryChipView,
  DropdownMenu,
  ErrorState,
  IconButton,
  OverdueBadge,
  PriorityBadge,
  Skeleton,
  StatusBadge,
  useToast,
} from '@/ui';
import {
  AddCostDialog,
  AssignDialog,
  CompleteDialog,
  LogTimeDialog,
  NoteDialog,
  SubWorkOrderDialog,
  WatchersDialog,
} from './dialogs';
import { WorkOrderForm } from './WorkOrderForm';

type DialogKey =
  'complete' | 'hold' | 'cancel' | 'assign' | 'watchers' | 'time' | 'cost' | 'sub' | 'edit' | null;

export function WorkOrderDetail({
  id,
  onClose,
  onNavigate,
  showBack,
  currentUser,
  onRefreshRequest,
}: {
  id: string;
  onClose: () => void;
  onNavigate: (id: string) => void;
  showBack: boolean;
  currentUser: UserRef;
  onDeleted?: () => void;
  onRefreshRequest?: () => void;
}) {
  const query = useWorkOrder(id);
  const client = useQueryClient();
  const invalidate = useInvalidate();
  const toast = useToast();
  const [dialog, setDialog] = useState<DialogKey>(null);
  const [busy, setBusy] = useState<string | null>(null);

  const refresh = (wo?: WorkOrderDetailType) => {
    if (wo)
      client.setQueryData(keys.workOrder(id), (prev: WorkOrderDetailType | undefined) => ({
        ...(prev || wo),
        ...wo,
        comments: wo.comments ?? prev?.comments,
        files: wo.files ?? prev?.files,
      }));
    invalidate(
      keys.workOrder(id),
      ['work-orders'],
      ['project-health'],
      ['project'],
      keys.notifications,
      keys.session,
    );
    onRefreshRequest?.();
  };

  const act = async (action: 'start' | 'resume' | 'open' | 'duplicate', label: string) => {
    setBusy(action);
    try {
      const result = await api<WorkOrderDetailType>(`/work-orders/${id}/${action}`, {
        method: 'POST',
        body: {},
      });
      toast.success(label);
      if (action === 'duplicate') {
        invalidate(['work-orders']);
        onNavigate(result.id);
      } else refresh(result);
    } catch (err) {
      toast.error(`Could not ${action}`, err instanceof ApiError ? err.message : undefined);
    } finally {
      setBusy(null);
    }
  };

  const patch = async (body: Record<string, unknown>, label: string) => {
    setBusy('patch');
    try {
      const result = await api<WorkOrderDetailType>(`/work-orders/${id}`, { method: 'PATCH', body });
      toast.success(label);
      refresh(result);
    } catch (err) {
      toast.error('Could not update', err instanceof ApiError ? err.message : undefined);
    } finally {
      setBusy(null);
    }
  };

  if (query.isPending)
    return (
      <div className="detail-panel">
        <Skeleton lines={8} />
      </div>
    );
  if (query.error || !query.data)
    return (
      <div className="detail-panel" style={{ padding: 16 }}>
        <ErrorState
          error={query.error}
          onRetry={() => query.refetch()}
          title={
            query.error instanceof ApiError && query.error.status === 404 ? 'Work order not found' : undefined
          }
        />
        <div style={{ marginTop: 12 }}>
          <Button size="sm" onClick={onClose}>
            Back to list
          </Button>
        </div>
      </div>
    );

  const wo = query.data;
  const p = wo.permissions;
  const t = new Set(wo.allowed_transitions);
  const closed = ['DONE', 'CANCELED', 'SKIPPED'].includes(wo.status);

  const primary = (() => {
    if (wo.status === 'DRAFT' && p.edit)
      return (
        <Button
          variant="primary"
          icon={<Play />}
          onClick={() => act('open', 'Work order opened')}
          loading={busy === 'open'}
          data-testid="action-open"
        >
          Open
        </Button>
      );
    if (wo.status === 'OPEN' && t.has('IN_PROGRESS') && p.start)
      return (
        <Button
          variant="primary"
          icon={<Play />}
          onClick={() => act('start', 'Work started')}
          loading={busy === 'start'}
          data-testid="action-start"
        >
          Start
        </Button>
      );
    if (wo.status === 'IN_PROGRESS' && t.has('DONE') && p.complete)
      return (
        <Button
          variant="success"
          icon={<CheckCircle2 />}
          onClick={() => setDialog('complete')}
          data-testid="action-complete"
        >
          Complete
        </Button>
      );
    if (wo.status === 'ON_HOLD' && t.has('IN_PROGRESS') && p.start)
      return (
        <Button
          variant="primary"
          icon={<Play />}
          onClick={() => act('resume', 'Work resumed')}
          loading={busy === 'resume'}
          data-testid="action-resume"
        >
          Resume
        </Button>
      );
    if (closed)
      return (
        <Button
          icon={<Copy />}
          onClick={() => act('duplicate', 'Work order duplicated')}
          loading={busy === 'duplicate'}
        >
          Duplicate
        </Button>
      );
    return null;
  })();

  const secondary = [];
  if (wo.status === 'OPEN' && t.has('DONE') && p.complete)
    secondary.push(
      <Button
        key="complete"
        variant="secondary"
        icon={<CheckCircle2 />}
        onClick={() => setDialog('complete')}
        data-testid="action-complete-open"
      >
        Complete
      </Button>,
    );
  if (wo.status === 'IN_PROGRESS' && t.has('ON_HOLD') && p.start)
    secondary.push(
      <Button
        key="hold"
        variant="secondary"
        icon={<Pause />}
        onClick={() => setDialog('hold')}
        data-testid="action-hold"
      >
        Hold
      </Button>,
    );
  if (p.assign)
    secondary.push(
      <Button
        key="assign"
        variant="secondary"
        icon={<UserPlus />}
        onClick={() => setDialog('assign')}
        data-testid="action-assign"
      >
        Assign
      </Button>,
    );

  const menu = [
    ...(p.edit
      ? [{ key: 'edit', label: 'Edit details', icon: <Pencil />, onSelect: () => setDialog('edit') }]
      : []),
    ...(p.edit && !closed
      ? [
          {
            key: 'blocked',
            label: wo.is_blocked ? 'Mark as unblocked' : 'Mark as blocked',
            icon: <ShieldAlert />,
            onSelect: () =>
              patch({ is_blocked: !wo.is_blocked }, wo.is_blocked ? 'Unblocked' : 'Marked as blocked'),
          },
        ]
      : []),
    ...(p.edit
      ? [{ key: 'watchers', label: 'Watchers', icon: <Eye />, onSelect: () => setDialog('watchers') }]
      : []),
    ...(!closed
      ? [
          {
            key: 'duplicate',
            label: 'Duplicate',
            icon: <Copy />,
            onSelect: () => act('duplicate', 'Work order duplicated'),
          },
        ]
      : []),
    {
      key: 'link',
      label: 'Copy link',
      icon: <Link2 />,
      onSelect: () =>
        navigator.clipboard
          ?.writeText(`${window.location.origin}/app/work-orders/${wo.id}`)
          .then(() => toast.success('Link copied')),
    },
    ...(t.has('CANCELED') && p.cancel
      ? [
          {
            key: 'cancel',
            label: 'Cancel work order',
            icon: <Ban />,
            danger: true,
            separatorBefore: true,
            onSelect: () => setDialog('cancel'),
          },
        ]
      : []),
  ];

  return (
    <article className="detail-panel" aria-label={`Work order #${wo.number}`} data-testid="work-order-detail">
      <div className="detail-head">
        <div className="row-between">
          <div className="row" style={{ gap: 8 }}>
            {showBack && (
              <IconButton label="Back to list" onClick={onClose} data-testid="detail-back">
                <ArrowLeft />
              </IconButton>
            )}
            <span className="mono text-muted">#{wo.number}</span>
            <StatusBadge status={wo.status} />
            <PriorityBadge priority={wo.priority} />
            {wo.is_overdue && <OverdueBadge />}
            {wo.is_blocked && <BlockedBadge />}
            {wo.recurrence && (
              <span className="badge badge-neutral" title="Repeats">
                <Repeat /> Every {wo.recurrence.interval}{' '}
                {wo.recurrence.frequency
                  .replace('ly', wo.recurrence.interval === 1 ? '' : 's')
                  .replace('dai', 'day')}
              </span>
            )}
          </div>
          {!showBack && (
            <IconButton label="Close details" onClick={onClose} data-testid="detail-close">
              <X />
            </IconButton>
          )}
        </div>
        <h2 className="detail-title" data-testid="detail-title">
          {wo.title}
        </h2>
        {wo.parent && (
          <div className="text-label text-muted row" style={{ gap: 6 }}>
            <GitBranch size={14} /> Sub-work order of{' '}
            <button type="button" className="link-button" onClick={() => onNavigate(wo.parent!.id)}>
              #{wo.parent.number} {wo.parent.title}
            </button>
          </div>
        )}
        <div className="detail-actions">
          {primary}
          {secondary}
          {menu.length > 0 && (
            <DropdownMenu
              label="More actions"
              items={menu}
              trigger={
                <button
                  type="button"
                  className="btn btn-secondary btn-icon"
                  aria-label="More actions"
                  data-testid="detail-menu"
                >
                  <MoreHorizontal />
                </button>
              }
            />
          )}
        </div>
      </div>

      {(wo.completion_note || wo.cancel_reason) && (
        <section className="detail-section">
          <div className={`notice ${wo.cancel_reason ? 'notice-warning' : ''}`}>
            {wo.cancel_reason ? <Ban /> : <CheckCircle2 />}
            <div>
              <strong>{wo.cancel_reason ? 'Canceled' : 'Completed'}</strong>{' '}
              {fmtDate(wo.canceled_at || wo.completed_at, true)}
              <div>{wo.cancel_reason || wo.completion_note}</div>
            </div>
          </div>
        </section>
      )}

      <section className="detail-section">
        <div className="detail-grid">
          <Field label="Assignees">
            <div className="row" style={{ gap: 8, flexWrap: 'wrap' }}>
              {wo.assignees.map((u) => (
                <span key={u.id} className="row" style={{ gap: 6 }}>
                  <Avatar user={u} size="sm" /> {u.name}
                </span>
              ))}
              {wo.assignee_teams.map((team) => (
                <span key={team.id} className="chip">
                  {team.name} (team)
                </span>
              ))}
              {wo.assignees.length === 0 && wo.assignee_teams.length === 0 && (
                <span className="text-muted">Unassigned</span>
              )}
            </div>
          </Field>
          <Field label="Due">
            <span style={wo.is_overdue ? { color: 'var(--color-danger)', fontWeight: 600 } : undefined}>
              {fmtDue(wo.due_at)}
            </span>
            {wo.start_at && (
              <div className="text-caption text-muted">Starts {fmtDate(wo.start_at, true)}</div>
            )}
          </Field>
          <Field label="Project">
            {wo.project ? <Link to={`/projects/${wo.project.id}`}>{wo.project.name}</Link> : '—'}
          </Field>
          <Field label="Team">{wo.team?.name || '—'}</Field>
          <Field label="Asset">
            {wo.primary_asset ? (
              <Link to={`/assets?asset=${wo.primary_asset.id}`}>{wo.primary_asset.name}</Link>
            ) : (
              '—'
            )}
            {wo.related_assets.length > 0 && (
              <div className="text-caption text-muted">
                + {wo.related_assets.map((a) => a.name).join(', ')}
              </div>
            )}
          </Field>
          <Field label="Location">{wo.location?.name || '—'}</Field>
          <Field label="Type">{humanize(wo.work_type)}</Field>
          <Field label="Estimate / actual">
            {fmtMinutes(wo.estimated_minutes)} / {fmtMinutes(wo.actual_minutes)}
          </Field>
          {wo.categories.length > 0 && (
            <Field label="Categories">
              <div className="row" style={{ flexWrap: 'wrap' }}>
                {wo.categories.map((c) => (
                  <CategoryChipView key={c.id} category={c} />
                ))}
              </div>
            </Field>
          )}
          {wo.budget_code && <Field label="Budget code">{wo.budget_code}</Field>}
          <Field label="Created">
            {wo.creator?.name || 'Unknown'} · {fmtDate(wo.created_at, true)}
          </Field>
          {wo.watchers.length > 0 && (
            <Field label="Watchers">
              <AvatarStack users={wo.watchers} max={5} />
            </Field>
          )}
        </div>
        {wo.description && (
          <div style={{ marginTop: 16 }}>
            <div className="detail-field-label">Description</div>
            <div className="description-block" data-testid="detail-description">
              {wo.description}
            </div>
          </div>
        )}
      </section>

      {(wo.children.length > 0 || (p.edit && !closed)) && (
        <section className="detail-section" data-testid="sub-work-orders">
          <div className="detail-section-title">
            <span>
              Sub-work orders{' '}
              <span className="count">
                {wo.child_progress.done}/{wo.child_progress.total} done
              </span>
            </span>
            {p.edit && !closed && (
              <Button size="sm" icon={<Plus />} onClick={() => setDialog('sub')} data-testid="add-sub">
                Add
              </Button>
            )}
          </div>
          {wo.children.length > 0 && (
            <div
              className="progress"
              style={{ marginBottom: 10 }}
              aria-label={`${wo.child_progress.percent}% of sub-work orders done`}
              role="progressbar"
              aria-valuenow={wo.child_progress.percent}
              aria-valuemin={0}
              aria-valuemax={100}
            >
              <span style={{ width: `${wo.child_progress.percent}%` }} />
            </div>
          )}
          {wo.children.length === 0 && (
            <div className="text-muted text-label">
              Break this work order into smaller steps for different people.
            </div>
          )}
          {wo.children.map((c: WorkOrderListItem) => (
            <button
              key={c.id}
              type="button"
              className="list-row"
              style={{ borderLeft: 'none', paddingLeft: 8, paddingRight: 8 }}
              onClick={() => onNavigate(c.id)}
              data-testid="sub-row"
            >
              <div className="list-row-main">
                <div className="list-row-title">
                  <span className="num">#{c.number}</span>
                  <span className="truncate">{c.title}</span>
                </div>
                <div className="list-row-meta">{fmtDue(c.due_at)}</div>
              </div>
              <div className="list-row-side">
                <StatusBadge status={c.status} />
                <AvatarStack users={c.assignees} />
              </div>
            </button>
          ))}
          {wo.parent_completion_policy === 'auto' && (
            <div className="text-caption text-muted" style={{ marginTop: 6 }}>
              This work order completes automatically when every sub-work order is done.
            </div>
          )}
        </section>
      )}

      <section className="detail-section" data-testid="comments-section">
        <div className="detail-section-title">
          <span>
            Comments <span className="count">{(wo.comments || []).filter((c) => !c.deleted).length}</span>
          </span>
        </div>
        <CommentThread
          comments={wo.comments || []}
          currentUser={currentUser}
          canComment={p.comment && !closed}
          onSubmit={async (body) => {
            try {
              const comment = await api<Comment>(`/work-orders/${wo.id}/comments`, {
                method: 'POST',
                body: { body },
              });
              client.setQueryData(keys.workOrder(id), (prev: WorkOrderDetailType | undefined) =>
                prev
                  ? {
                      ...prev,
                      comments: [...(prev.comments || []), comment],
                      comment_count: prev.comment_count + 1,
                    }
                  : prev,
              );
              invalidate(['work-orders']);
            } catch (err) {
              toast.error('Could not post comment', err instanceof ApiError ? err.message : undefined);
              throw err;
            }
          }}
          onEdit={async (comment, body) => {
            try {
              const updated = await api<Comment>(`/comments/${comment.id}`, {
                method: 'PATCH',
                body: { body },
              });
              client.setQueryData(keys.workOrder(id), (prev: WorkOrderDetailType | undefined) =>
                prev
                  ? {
                      ...prev,
                      comments: (prev.comments || []).map((c) => (c.id === updated.id ? updated : c)),
                    }
                  : prev,
              );
            } catch (err) {
              toast.error('Could not edit comment', err instanceof ApiError ? err.message : undefined);
            }
          }}
          onDelete={async (comment) => {
            try {
              await api(`/comments/${comment.id}`, { method: 'DELETE' });
              invalidate(keys.workOrder(id), ['work-orders']);
            } catch (err) {
              toast.error('Could not delete comment', err instanceof ApiError ? err.message : undefined);
            }
          }}
        />
      </section>

      <section className="detail-section" data-testid="files-section">
        <div className="detail-section-title">
          <span>
            Files <span className="count">{(wo.files || []).length}</span>
          </span>
        </div>
        <div className="stack">
          <AttachmentList
            files={wo.files || []}
            currentUser={currentUser}
            canDelete={(f) => p.edit || f.uploaded_by?.id === currentUser.id}
            onChanged={() => invalidate(keys.workOrder(id))}
          />
          {p.upload && !closed && (
            <AttachmentUploader
              entityType="work_order"
              entityId={wo.id}
              onUploaded={() => invalidate(keys.workOrder(id))}
              compact
            />
          )}
        </div>
      </section>

      <section className="detail-section" data-testid="time-cost-section">
        <div className="detail-section-title">
          <span>
            Time &amp; costs{' '}
            <span className="count">
              {fmtMinutes(wo.actual_minutes)} · {fmtMoney(wo.total_cost)}
            </span>
          </span>
          <span className="inline-actions">
            {p.log_time && !closed && (
              <Button size="sm" icon={<Timer />} onClick={() => setDialog('time')} data-testid="log-time">
                Log time
              </Button>
            )}
            {p.log_cost && !closed && (
              <Button
                size="sm"
                icon={<DollarSign />}
                onClick={() => setDialog('cost')}
                data-testid="add-cost"
              >
                Add cost
              </Button>
            )}
          </span>
        </div>
        {wo.time_entries.length === 0 && wo.cost_entries.length === 0 && (
          <div className="text-muted text-label">Nothing logged yet.</div>
        )}
        {wo.time_entries.length > 0 && (
          <table className="table" style={{ marginBottom: wo.cost_entries.length ? 12 : 0 }}>
            <thead>
              <tr>
                <th>Who</th>
                <th>Time</th>
                <th>Note</th>
                <th>When</th>
              </tr>
            </thead>
            <tbody>
              {wo.time_entries.map((entry) => (
                <tr key={entry.id}>
                  <td>{entry.user.name}</td>
                  <td>{fmtMinutes(entry.minutes)}</td>
                  <td>{entry.note || '—'}</td>
                  <td>{fmtDate(entry.created_at, true)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
        {wo.cost_entries.length > 0 && (
          <table className="table">
            <thead>
              <tr>
                <th>Type</th>
                <th>Description</th>
                <th className="num">Amount</th>
              </tr>
            </thead>
            <tbody>
              {wo.cost_entries.map((entry) => (
                <tr key={entry.id}>
                  <td>{humanize(entry.type)}</td>
                  <td>{entry.description || '—'}</td>
                  <td className="num">{fmtMoney(entry.amount)}</td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </section>

      <HistorySection id={wo.id} statusHistory={wo.status_history} />

      {dialog === 'complete' && (
        <CompleteDialog
          wo={wo}
          open
          onClose={() => setDialog(null)}
          onDone={(r) => {
            setDialog(null);
            refresh(r);
          }}
        />
      )}
      {dialog === 'hold' && (
        <NoteDialog
          wo={wo}
          open
          onClose={() => setDialog(null)}
          onDone={(r) => {
            setDialog(null);
            refresh(r);
          }}
          action="hold"
          title="Put on hold"
          label="Why is this on hold?"
          confirmLabel="Put on hold"
        />
      )}
      {dialog === 'cancel' && (
        <NoteDialog
          wo={wo}
          open
          onClose={() => setDialog(null)}
          onDone={(r) => {
            setDialog(null);
            refresh(r);
          }}
          action="cancel"
          title={`Cancel #${wo.number}?`}
          label="Reason"
          confirmLabel="Cancel work order"
          danger
          required
        />
      )}
      {dialog === 'assign' && (
        <AssignDialog
          wo={wo}
          open
          onClose={() => setDialog(null)}
          onDone={(r) => {
            setDialog(null);
            refresh(r);
          }}
        />
      )}
      {dialog === 'watchers' && (
        <WatchersDialog
          wo={wo}
          open
          onClose={() => setDialog(null)}
          onDone={(r) => {
            setDialog(null);
            refresh(r);
          }}
        />
      )}
      {dialog === 'time' && (
        <LogTimeDialog
          wo={wo}
          open
          onClose={() => setDialog(null)}
          onDone={() => {
            setDialog(null);
            refresh();
          }}
        />
      )}
      {dialog === 'cost' && (
        <AddCostDialog
          wo={wo}
          open
          onClose={() => setDialog(null)}
          onDone={() => {
            setDialog(null);
            refresh();
          }}
        />
      )}
      {dialog === 'sub' && (
        <SubWorkOrderDialog
          parent={wo}
          open
          onClose={() => setDialog(null)}
          onDone={() => {
            setDialog(null);
            refresh();
          }}
        />
      )}
      {dialog === 'edit' && (
        <WorkOrderForm
          mode="edit"
          workOrder={wo}
          onClose={() => setDialog(null)}
          onSaved={(r) => {
            setDialog(null);
            refresh(r);
          }}
        />
      )}
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

function HistorySection({ id, statusHistory }: { id: string; statusHistory: StatusHistoryEntry[] }) {
  const [expanded, setExpanded] = useState(false);
  const history = useQuery<{ status_history: StatusHistoryEntry[]; events: AuditEvent[] }>({
    queryKey: ['work-order-history', id],
    queryFn: () => api(`/work-orders/${id}/history`),
    enabled: expanded,
  });
  const items =
    expanded && history.data ? auditItems(history.data.events) : statusHistoryItems(statusHistory);
  return (
    <section className="detail-section" data-testid="history-section">
      <div className="detail-section-title">
        <span>History</span>
        <button type="button" className="link-button text-caption" onClick={() => setExpanded((v) => !v)}>
          {expanded ? 'Status changes only' : 'Show full audit trail'}
        </button>
      </div>
      {expanded && history.isPending ? <Skeleton lines={3} /> : <ActivityTimeline items={items} />}
      {!expanded && statusHistory.length > 0 && (
        <div className="text-caption text-muted" style={{ marginTop: 8 }}>
          Current status: {STATUS_LABEL[statusHistory[statusHistory.length - 1]?.to_status] || ''}
        </div>
      )}
    </section>
  );
}
