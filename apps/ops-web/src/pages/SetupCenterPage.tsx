import { useState } from 'react';
import { Link } from 'react-router-dom';
import { ArrowRight, BookOpen, Check, Clock } from 'lucide-react';
import { api } from '@/api/client';
import { keys, useApiMutation, useSetup } from '@/api/hooks';
import { useCurrentSession } from '@/auth/SessionProvider';
import type { SetupTask } from '@/contracts/types';
import { Button, ErrorState, PageHeader, Skeleton, Tabs } from '@/ui';
import { cn } from '@/lib/cn';

const route = (r: string) => r.replace(/^\/app/, '');

export function SetupCenterPage() {
  const { session, refresh } = useCurrentSession();
  const setup = useSetup();
  const [tab, setTab] = useState<'checklist' | 'guide'>(
    window.location.hash === '#guide' ? 'guide' : 'checklist',
  );
  const banner = useApiMutation(
    (dismissed: boolean) => api('/ops/setup/banner', { method: 'POST', body: { dismissed } }),
    [keys.setup],
  );
  const guideRead = useApiMutation(() => api('/ops/setup/guide-read', { method: 'POST' }), [keys.setup]);

  if (setup.isPending) return <Skeleton lines={6} />;
  if (setup.error || !setup.data) return <ErrorState error={setup.error} onRetry={() => setup.refetch()} />;
  const data = setup.data;
  const ring = `conic-gradient(var(--color-success) ${data.percent * 3.6}deg, var(--color-bg-subtle) 0)`;

  return (
    <div className="page">
      <PageHeader
        title="Setup Center"
        subtitle={`Get ${session.organization.name} ready to run its operations.`}
        testId="setup-center"
      />
      <div className="page-body stack" style={{ gap: 20 }}>
        <div className="setup-hero">
          <div className="stack" style={{ gap: 8 }}>
            <div className="text-label text-muted">Overall progress</div>
            <div style={{ fontSize: 22, fontWeight: 700 }}>
              {data.complete
                ? 'Setup complete'
                : `${data.steps_left} required step${data.steps_left === 1 ? '' : 's'} left`}
            </div>
            <div className="text-secondary">
              {data.required_done} of {data.required_total} required tasks done. Progress is calculated from
              what is actually configured, not from checkboxes.
            </div>
            <div className="row" style={{ marginTop: 6, gap: 10 }}>
              {data.next_step && (
                <Link to={route(data.next_step.route)} className="btn btn-primary" data-testid="setup-next">
                  Next: {data.next_step.title} <ArrowRight size={16} />
                </Link>
              )}
              <Button
                variant="ghost"
                onClick={async () => {
                  await banner.mutateAsync(!(data.banner_dismissed || session.setup_banner_dismissed));
                  await refresh();
                }}
              >
                {data.banner_dismissed || session.setup_banner_dismissed
                  ? 'Show setup banner'
                  : 'Hide setup banner'}
              </Button>
            </div>
          </div>
          <div
            className="setup-ring"
            style={{ background: ring }}
            aria-label={`${data.percent}% complete`}
            role="img"
          >
            <span>{data.percent}%</span>
          </div>
        </div>

        <Tabs
          label="Setup sections"
          value={tab}
          onChange={setTab}
          tabs={[
            { key: 'checklist', label: 'Checklist' },
            { key: 'guide', label: 'Officer guide' },
          ]}
        />

        {tab === 'checklist' &&
          data.phases.map((phase, index) => (
            <section key={phase.key} className="setup-phase" aria-labelledby={`phase-${phase.key}`}>
              <div className="setup-phase-head">
                <span className={cn('setup-phase-num', phase.complete && 'done')}>
                  {phase.complete ? <Check size={16} /> : index + 1}
                </span>
                <div style={{ flex: 1 }}>
                  <h2 id={`phase-${phase.key}`} className="card-title">
                    {phase.title}
                  </h2>
                  <div className="text-muted text-label" style={{ fontWeight: 450 }}>
                    {phase.description}
                  </div>
                </div>
                <div className="text-caption text-muted row" style={{ gap: 6 }}>
                  <Clock size={12} /> ~{phase.estimated_minutes} min · {phase.required_done}/
                  {phase.required_total} required
                </div>
              </div>
              {phase.tasks.map((task) => (
                <TaskRow key={task.key} task={task} onGuide={() => setTab('guide')} />
              ))}
            </section>
          ))}

        {tab === 'guide' && (
          <section className="card" id="guide">
            <div className="card-head">
              <span className="card-title row">
                <BookOpen size={18} /> Officer onboarding guide
              </span>
              <Button
                size="sm"
                variant="primary"
                icon={<Check />}
                onClick={() => guideRead.mutate(undefined as never)}
                disabled={data.phases[0].tasks.find((t) => t.key === 'officer_guide')?.complete}
                data-testid="guide-read"
              >
                {data.phases[0].tasks.find((t) => t.key === 'officer_guide')?.complete
                  ? 'Marked as read'
                  : 'Mark as read'}
              </Button>
            </div>
            <div className="card-body stack" style={{ gap: 14, maxWidth: 760, lineHeight: 1.6 }}>
              <p>
                <strong>Projects</strong> are the chapter's programmes: a competition build, showcase, or
                general operations. Each has a lead, milestones and a budget code.
              </p>
              <p>
                <strong>Work orders</strong> are the unit of work. Anything that needs doing, from a wheel-hub
                inspection to booking a showcase table, is a work order with an assignee, a due date and a
                status that moves from Open to In progress to Done. Large tasks break into sub-work orders.
              </p>
              <p>
                <strong>Assets</strong> are the things you maintain: the rover and its subsystems, printers,
                test equipment. Attaching an asset to a work order builds its maintenance history
                automatically.
              </p>
              <p>
                <strong>Teams</strong> route work. Assign a work order to a team and every member sees it;
                team leads can reassign within their team. Roles decide what people can do; teams decide what
                they see.
              </p>
              <p>
                <strong>Locations and categories</strong> are the labels reports group by. Keep them short and
                consistent. The defaults match a typical student-chapter shop.
              </p>
              <p className="text-muted">
                Requests, parts inventory, procedures, maintenance plans and automations are planned for later
                releases and do not appear in navigation until they ship.
              </p>
            </div>
          </section>
        )}
      </div>
    </div>
  );
}

function TaskRow({ task, onGuide }: { task: SetupTask; onGuide: () => void }) {
  const done = task.complete;
  return (
    <div
      className={cn('setup-task', !task.available && 'unavailable')}
      data-testid={`setup-task-${task.key}`}
    >
      <span className={cn('setup-task-check', done && 'done')} aria-hidden="true">
        {done && <Check />}
      </span>
      <div style={{ minWidth: 0 }}>
        <div className="setup-task-title">
          {task.title}
          {!task.required && (
            <span className="badge badge-neutral" style={{ marginLeft: 8 }}>
              Optional
            </span>
          )}
        </div>
        <div className="setup-task-desc">
          {task.description}
          {task.available && task.target > 1 && (
            <>
              {' '}
              · {task.current}/{task.target}
            </>
          )}
        </div>
      </div>
      <div>
        {!task.available ? (
          <span className="text-caption text-muted">Not in this release</span>
        ) : done ? (
          <span className="text-caption" style={{ color: 'var(--color-success)', fontWeight: 600 }}>
            Done
          </span>
        ) : task.key === 'officer_guide' ? (
          <Button size="sm" onClick={onGuide}>
            Read guide
          </Button>
        ) : (
          <Link to={route(task.route)} className="btn btn-secondary btn-sm">
            {task.minutes} min <ArrowRight size={14} />
          </Link>
        )}
      </div>
    </div>
  );
}
