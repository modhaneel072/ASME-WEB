import { useMemo } from 'react';
import { Link, useNavigate } from 'react-router-dom';
import { AlertTriangle, ArrowUpDown, ChevronDown, FolderKanban, Plus } from 'lucide-react';
import { useCurrentSession } from '@/auth/SessionProvider';
import { useLookups, useProjects } from '@/api/hooks';
import type { ProjectListItem } from '@/contracts/types';
import { useListParams } from '@/lib/listParams';
import { fmtDate, humanize } from '@/lib/format';
import { userFilterOptions } from '@/lib/options';
import {
  Avatar,
  Badge,
  Button,
  DropdownMenu,
  EmptyState,
  ErrorState,
  FilterChip,
  PageHeader,
  SearchField,
  Segmented,
  Skeleton,
} from '@/ui';
import { ProjectForm } from './ProjectForm';

const SORTS = [
  { value: 'name_asc', label: 'Name' },
  { value: 'updated_desc', label: 'Recently updated' },
  { value: 'target_asc', label: 'Target date' },
  { value: 'risk_desc', label: 'Risk' },
];

const STATUS_TONE: Record<string, string> = {
  planning: 'info',
  active: 'success',
  on_hold: 'warning',
  completed: 'neutral',
  archived: 'neutral',
};
const RISK_TONE: Record<string, string> = { low: 'success', medium: 'warning', high: 'critical' };

export function ProjectsPage() {
  const { can } = useCurrentSession();
  const lookups = useLookups();
  const navigate = useNavigate();
  const list = useListParams(['view', 'new']);
  const view = (list.state.extras.view as 'active' | 'all' | 'archived') || 'active';
  const creating = list.state.extras.new === '1';
  const params = useMemo(() => ({ ...list.apiParams, view }), [list.apiParams, view]);
  const query = useProjects(params);
  const items = query.data?.items || [];
  const filterValue = (key: string) => list.state.filters[key] || [];
  const setFilter = (key: string) => (values: string[]) => list.update({ filters: { [key]: values } });
  const canCreate = can('project.create');
  const sort = list.state.sort || 'name_asc';

  return (
    <div className="page">
      <PageHeader
        title="Projects"
        subtitle="Competition builds, showcases and chapter operations."
        testId="projects-page"
        actions={
          canCreate && (
            <Button
              variant="primary"
              icon={<Plus />}
              onClick={() => list.update({ extras: { new: '1' } })}
              data-testid="new-project"
            >
              New project
            </Button>
          )
        }
      />
      <div className="page-toolbar">
        <Segmented
          label="Project view"
          value={view}
          onChange={(v) => list.update({ extras: { view: v === 'active' ? null : v } })}
          options={[
            { value: 'active', label: 'Active' },
            { value: 'all', label: 'All' },
            { value: 'archived', label: 'Archived' },
          ]}
        />
        <SearchField
          value={list.state.q}
          onChange={(q) => list.update({ q })}
          placeholder="Search name or code"
        />
        <div className="filter-bar" style={{ marginLeft: 'auto' }}>
          <FilterChip
            label="Status"
            options={['planning', 'active', 'on_hold', 'completed', 'archived'].map((s) => ({
              value: s,
              label: humanize(s),
            }))}
            value={filterValue('status')}
            onChange={setFilter('status')}
          />
          <FilterChip
            label="Risk"
            options={['low', 'medium', 'high'].map((s) => ({ value: s, label: humanize(s) }))}
            value={filterValue('risk')}
            onChange={setFilter('risk')}
          />
          <FilterChip
            label="Lead"
            options={userFilterOptions(lookups.data, false)}
            value={filterValue('lead')}
            onChange={setFilter('lead')}
          />
          {list.activeFilterCount > 0 && (
            <button type="button" className="filter-chip-clear" onClick={list.clearFilters}>
              Clear all
            </button>
          )}
          <DropdownMenu
            label="Sort projects"
            trigger={
              <button type="button" className="filter-chip">
                <ArrowUpDown /> {SORTS.find((s) => s.value === sort)?.label} <ChevronDown />
              </button>
            }
            items={SORTS.map((s) => ({
              key: s.value,
              label: s.label,
              onSelect: () => list.update({ sort: s.value === 'name_asc' ? null : s.value }),
            }))}
          />
        </div>
      </div>
      <div className="page-body">
        {query.isPending ? (
          <Skeleton lines={5} />
        ) : query.error ? (
          <ErrorState error={query.error} onRetry={() => query.refetch()} />
        ) : items.length === 0 ? (
          <div className="card">
            <EmptyState
              icon={<FolderKanban />}
              title={
                list.activeFilterCount
                  ? 'No projects match'
                  : view === 'archived'
                    ? 'No archived projects'
                    : 'No projects yet'
              }
              body={
                list.activeFilterCount
                  ? 'Try clearing a filter.'
                  : canCreate
                    ? 'Create the first project, assign a lead and add milestones.'
                    : 'Projects will appear here once an officer creates them.'
              }
              actions={
                canCreate && !list.activeFilterCount && view !== 'archived' ? (
                  <Button
                    variant="primary"
                    icon={<Plus />}
                    onClick={() => list.update({ extras: { new: '1' } })}
                  >
                    New project
                  </Button>
                ) : list.activeFilterCount ? (
                  <Button onClick={list.clearFilters}>Clear filters</Button>
                ) : undefined
              }
            />
          </div>
        ) : (
          <div className="project-grid" data-testid="project-grid">
            {items.map((p) => (
              <ProjectCard key={p.id} project={p} />
            ))}
          </div>
        )}
        {query.data?.next_cursor && (
          <div className="text-caption text-muted" style={{ marginTop: 12 }}>
            Showing the first {items.length} projects. Narrow the search to find others.
          </div>
        )}
      </div>
      {creating && (
        <ProjectForm
          onClose={() => list.update({ extras: { new: null } })}
          onSaved={(p) => {
            list.update({ extras: { new: null } });
            navigate(`/projects/${p.id}`);
          }}
        />
      )}
    </div>
  );
}

function ProjectCard({ project }: { project: ProjectListItem }) {
  return (
    <Link to={`/projects/${project.id}`} className="project-card" data-testid="project-card">
      <div className="row-between" style={{ alignItems: 'flex-start' }}>
        <div style={{ minWidth: 0 }}>
          <div className="project-card-title truncate">{project.name}</div>
          <div className="project-card-code">
            {project.code}
            {project.academic_year ? ` · ${project.academic_year}` : ''}
          </div>
        </div>
        <div className="row" style={{ gap: 6 }}>
          <Badge tone={STATUS_TONE[project.status] || 'neutral'}>{humanize(project.status)}</Badge>
          <Badge
            tone={RISK_TONE[project.risk_level] || 'neutral'}
            icon={project.risk_level === 'high' ? <AlertTriangle /> : undefined}
          >
            {humanize(project.risk_level)} risk
          </Badge>
        </div>
      </div>
      <div>
        <div className="row-between text-caption text-muted" style={{ marginBottom: 4 }}>
          <span>Completion</span>
          <span>{project.completion_percent}%</span>
        </div>
        <div
          className="progress"
          role="progressbar"
          aria-valuenow={project.completion_percent}
          aria-valuemin={0}
          aria-valuemax={100}
          aria-label="Completion"
        >
          <span style={{ width: `${project.completion_percent}%` }} />
        </div>
      </div>
      <div className="project-card-stats">
        <div>
          <strong>{project.open_work_orders}</strong>open
        </div>
        <div>
          <strong style={project.overdue_work_orders ? { color: 'var(--color-danger)' } : undefined}>
            {project.overdue_work_orders}
          </strong>
          overdue
        </div>
        <div>
          <strong>{project.team_count}</strong>teams
        </div>
      </div>
      <div className="row-between text-caption text-muted">
        <span className="row" style={{ gap: 6 }}>
          {project.lead ? (
            <>
              <Avatar user={project.lead} size="sm" /> {project.lead.name}
            </>
          ) : (
            'No lead'
          )}
        </span>
        <span>
          {project.next_milestone
            ? `${project.next_milestone.name} · ${project.next_milestone.days_left}d`
            : project.target_date
              ? `Target ${fmtDate(project.target_date)}`
              : ''}
        </span>
      </div>
    </Link>
  );
}
