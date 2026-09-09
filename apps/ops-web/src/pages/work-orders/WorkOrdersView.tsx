import { useCallback, useMemo, useState } from 'react';
import { useSearchParams } from 'react-router-dom';
import {
  ArrowUpDown,
  Bookmark,
  ClipboardList,
  LayoutList,
  MessageSquare,
  Plus,
  Table2,
  GitBranch,
  ChevronDown,
  Repeat,
} from 'lucide-react';
import { useCurrentSession } from '@/auth/SessionProvider';
import { useLookups, useWorkOrdersInfinite } from '@/api/hooks';
import type { WorkOrderListItem } from '@/contracts/types';
import { useListParams } from '@/lib/listParams';
import { useIsCompact } from '@/lib/useMediaQuery';
import { fmtDue } from '@/lib/format';
import { cn } from '@/lib/cn';
import {
  assetFilterOptions,
  categoryFilterOptions,
  dueOptions,
  locationFilterOptions,
  priorityOptions,
  projectFilterOptions,
  statusOptions,
  teamFilterOptions,
  userFilterOptions,
  WO_STATUS_CLOSED,
  WO_STATUS_OPEN,
  workTypeOptions,
} from '@/lib/options';
import {
  AvatarStack,
  BlockedBadge,
  Button,
  CategoryChipView,
  DropdownMenu,
  EmptyState,
  ErrorState,
  FilterChip,
  OverdueBadge,
  PageHeader,
  PriorityBadge,
  SearchField,
  Skeleton,
  StatusBadge,
  Tabs,
} from '@/ui';
import { WorkOrderForm } from './WorkOrderForm';
import { WorkOrderDetail } from './WorkOrderDetail';
import { SavedViews } from './SavedViews';

const SORTS = [
  { value: 'due_asc', label: 'Due date' },
  { value: 'priority_desc', label: 'Priority' },
  { value: 'updated_desc', label: 'Recently updated' },
  { value: 'created_desc', label: 'Newest' },
  { value: 'number_desc', label: 'Number (high → low)' },
  { value: 'number_asc', label: 'Number (low → high)' },
  { value: 'project', label: 'Project' },
];

/**
 * Work-orders list + detail. Used standalone (/work-orders) and inside a project
 * (projectId set: the project filter is pinned and hidden).
 */
export function WorkOrdersView({ projectId, embedded = false }: { projectId?: string; embedded?: boolean }) {
  const { can, session } = useCurrentSession();
  const lookups = useLookups();
  const [params, setParams] = useSearchParams();
  const list = useListParams(['tab', 'wo', 'new', 'view', 'edit']);
  const compact = useIsCompact();

  const tab = (list.state.extras.tab as 'todo' | 'done') || 'todo';
  const selectedId = list.state.extras.wo || null;
  const creating = list.state.extras.new === '1';
  const viewMode = (list.state.extras.view as 'panel' | 'table') || 'panel';
  const sort = list.state.sort || 'due_asc';

  const apiParams = useMemo(
    () => ({ ...list.apiParams, tab, sort, ...(projectId ? { 'filter[project]': projectId } : {}) }),
    [list.apiParams, tab, sort, projectId],
  );
  const query = useWorkOrdersInfinite(apiParams);
  const items = useMemo(() => query.data?.pages.flatMap((p) => p.items) || [], [query.data]);
  const counts = query.data?.pages[0]?.counts;

  const select = useCallback(
    (id: string | null) => {
      const next = new URLSearchParams(params);
      if (id) next.set('wo', id);
      else next.delete('wo');
      next.delete('edit');
      setParams(next, { replace: false });
    },
    [params, setParams],
  );
  const setCreating = (open: boolean) => list.update({ extras: { new: open ? '1' : null } });

  const filterValue = (key: string) => list.state.filters[key] || [];
  const setFilter = (key: string) => (values: string[]) => list.update({ filters: { [key]: values } });

  const canCreate = can('work_order.create');
  const [detailKey, setDetailKey] = useState(0);
  const showDetail = !!selectedId;

  const header = (
    <>
      {!embedded && (
        <PageHeader
          title="Work Orders"
          subtitle={counts ? `${counts.todo} to do · ${counts.done} done` : undefined}
          testId="work-orders-page"
          actions={
            canCreate && (
              <Button
                variant="primary"
                icon={<Plus />}
                onClick={() => setCreating(true)}
                data-testid="new-work-order"
              >
                New work order
              </Button>
            )
          }
        />
      )}
      <div className="page-toolbar" style={embedded ? { paddingLeft: 0, paddingRight: 0 } : undefined}>
        <Tabs
          label="Work order status"
          value={tab}
          onChange={(key) =>
            list.update({
              extras: { tab: key === 'todo' ? null : key, wo: null },
              filters: { status: undefined },
            })
          }
          tabs={[
            { key: 'todo', label: 'To Do', count: counts?.todo ?? null },
            { key: 'done', label: 'Done', count: counts?.done ?? null },
          ]}
        />
        <div className="row" style={{ marginLeft: 'auto', gap: 8, flexWrap: 'wrap' }}>
          <SearchField
            value={list.state.q}
            onChange={(q) => list.update({ q })}
            placeholder="Search title, description or #number"
          />
          {embedded && canCreate && (
            <Button
              variant="primary"
              icon={<Plus />}
              onClick={() => setCreating(true)}
              data-testid="new-work-order"
            >
              New work order
            </Button>
          )}
        </div>
      </div>
      <div className="page-toolbar" style={embedded ? { paddingLeft: 0, paddingRight: 0 } : undefined}>
        <div className="filter-bar" data-testid="filter-bar">
          <FilterChip
            label="Status"
            options={statusOptions(tab === 'done' ? WO_STATUS_CLOSED : WO_STATUS_OPEN)}
            value={filterValue('status')}
            onChange={setFilter('status')}
          />
          <FilterChip
            label="Priority"
            options={priorityOptions}
            value={filterValue('priority')}
            onChange={setFilter('priority')}
          />
          <FilterChip
            label="Assignee"
            options={userFilterOptions(lookups.data)}
            value={filterValue('assignee')}
            onChange={setFilter('assignee')}
          />
          <FilterChip
            label="Team"
            options={teamFilterOptions(lookups.data)}
            value={filterValue('team')}
            onChange={setFilter('team')}
          />
          {!projectId && (
            <FilterChip
              label="Project"
              options={projectFilterOptions(lookups.data, true)}
              value={filterValue('project')}
              onChange={setFilter('project')}
            />
          )}
          <FilterChip
            label="Due"
            options={dueOptions}
            value={filterValue('due')}
            onChange={setFilter('due')}
            single
          />
          <FilterChip
            label="Location"
            options={locationFilterOptions(lookups.data)}
            value={filterValue('location')}
            onChange={setFilter('location')}
          />
          <FilterChip
            label="Asset"
            options={assetFilterOptions(lookups.data)}
            value={filterValue('asset')}
            onChange={setFilter('asset')}
          />
          <FilterChip
            label="Category"
            options={categoryFilterOptions(lookups.data)}
            value={filterValue('category')}
            onChange={setFilter('category')}
          />
          <FilterChip
            label="Type"
            options={workTypeOptions}
            value={filterValue('work_type')}
            onChange={setFilter('work_type')}
          />
          <FilterChip
            label="Blocked"
            options={[
              { value: '1', label: 'Blocked only' },
              { value: '0', label: 'Not blocked' },
            ]}
            value={filterValue('blocked')}
            onChange={setFilter('blocked')}
            single
          />
          {list.activeFilterCount > 0 && (
            <button
              type="button"
              className="filter-chip-clear"
              onClick={list.clearFilters}
              data-testid="clear-filters"
            >
              Clear all
            </button>
          )}
          <div className="filter-summary">
            <SavedViews
              entityType="work_order"
              filters={list.state.filters}
              sort={sort}
              viewType={viewMode}
              onApply={(f, s, v) =>
                list.update({
                  filters: Object.fromEntries(
                    Object.keys({ ...list.state.filters, ...f }).map((k) => [k, f[k]]),
                  ),
                  sort: s,
                  extras: { view: v === 'table' ? 'table' : null },
                })
              }
              canShare={can('saved_filter.share')}
              icon={<Bookmark size={14} />}
            />
            <DropdownMenu
              label="Sort"
              trigger={
                <button type="button" className="filter-chip" data-testid="sort-menu">
                  <ArrowUpDown /> {SORTS.find((s) => s.value === sort)?.label || 'Sort'} <ChevronDown />
                </button>
              }
              items={SORTS.map((s) => ({
                key: s.value,
                label: s.label,
                onSelect: () => list.update({ sort: s.value === 'due_asc' ? null : s.value }),
              }))}
            />
            <div className="segmented" role="group" aria-label="View">
              <button
                type="button"
                aria-pressed={viewMode === 'panel'}
                onClick={() => list.update({ extras: { view: null } })}
                aria-label="Panel view"
                title="Panel view"
              >
                <LayoutList size={16} />
              </button>
              <button
                type="button"
                aria-pressed={viewMode === 'table'}
                onClick={() => list.update({ extras: { view: 'table' } })}
                aria-label="Table view"
                title="Table view"
              >
                <Table2 size={16} />
              </button>
            </div>
          </div>
        </div>
      </div>
    </>
  );

  const listBody = query.isPending ? (
    <div className="list">
      <Skeleton lines={6} />
    </div>
  ) : query.error ? (
    <ErrorState error={query.error} onRetry={() => query.refetch()} />
  ) : items.length === 0 ? (
    <div className="list">
      <EmptyState
        icon={<ClipboardList />}
        title={
          list.activeFilterCount
            ? 'No work orders match these filters'
            : tab === 'done'
              ? 'Nothing completed yet'
              : 'No open work orders'
        }
        body={
          list.activeFilterCount
            ? 'Try clearing a filter or two.'
            : tab === 'done'
              ? 'Completed and canceled work orders will show up here.'
              : canCreate
                ? 'Create the first work order for this team or project.'
                : 'Work assigned to you or your teams will show up here.'
        }
        actions={
          list.activeFilterCount ? (
            <Button onClick={list.clearFilters}>Clear filters</Button>
          ) : (
            canCreate && (
              <Button variant="primary" icon={<Plus />} onClick={() => setCreating(true)}>
                New work order
              </Button>
            )
          )
        }
      />
    </div>
  ) : viewMode === 'table' && !(compact && showDetail) ? (
    <WorkOrderTable items={items} selectedId={selectedId} onSelect={select} />
  ) : (
    <div className="list" data-testid="work-order-list">
      {items.map((wo) => (
        <WorkOrderRow key={wo.id} wo={wo} selected={wo.id === selectedId} onSelect={() => select(wo.id)} />
      ))}
      {query.hasNextPage && (
        <div className="load-more">
          <Button onClick={() => query.fetchNextPage()} loading={query.isFetchingNextPage}>
            Load more
          </Button>
        </div>
      )}
    </div>
  );

  return (
    <div className="page">
      {header}
      <div className={cn('page-body', embedded && 'page-body-flush')}>
        <div className={cn('master-detail', !showDetail && 'no-detail', showDetail && 'has-detail')}>
          <div className="master">{listBody}</div>
          {showDetail && selectedId && (
            <div className="detail">
              <WorkOrderDetail
                key={`${selectedId}-${detailKey}`}
                id={selectedId}
                onClose={() => select(null)}
                onNavigate={(id) => select(id)}
                showBack={compact}
                onDeleted={() => select(null)}
                onRefreshRequest={() => setDetailKey((k) => k + 1)}
                currentUser={session.user}
              />
            </div>
          )}
        </div>
      </div>
      {creating && (
        <WorkOrderForm
          mode="create"
          defaults={projectId ? { project_id: projectId } : undefined}
          onClose={() => setCreating(false)}
          onCreated={(wo) => {
            // one URL update: close the pane and select the new work order
            list.update({ extras: { new: null, wo: wo.id, edit: null } }, false);
          }}
        />
      )}
    </div>
  );
}

function WorkOrderRow({
  wo,
  selected,
  onSelect,
}: {
  wo: WorkOrderListItem;
  selected: boolean;
  onSelect: () => void;
}) {
  return (
    <button
      type="button"
      className="list-row"
      aria-current={selected || undefined}
      onClick={onSelect}
      data-testid="work-order-row"
      data-number={wo.number}
    >
      <div className="list-row-main">
        <div className="list-row-title">
          <span className="num">#{wo.number}</span>
          <span className="truncate">{wo.title}</span>
          {wo.parent_work_order_id && (
            <GitBranch size={13} className="text-muted" aria-label="Sub-work order" />
          )}
        </div>
        <div className="list-row-meta">
          {wo.project && <span>{wo.project.name}</span>}
          {wo.team && <span>· {wo.team.name}</span>}
          {wo.primary_asset && <span>· {wo.primary_asset.name}</span>}
          <span className={cn(wo.is_overdue && 'overdue')}>· {fmtDue(wo.due_at)}</span>
          {wo.comment_count > 0 && (
            <span className="row" style={{ gap: 3 }}>
              · <MessageSquare size={12} /> {wo.comment_count}
            </span>
          )}
          {wo.child_count > 0 && (
            <span className="row" style={{ gap: 3 }}>
              · <GitBranch size={12} /> {wo.child_count}
            </span>
          )}
          {wo.categories.slice(0, 2).map((c) => (
            <CategoryChipView key={c.id} category={c} />
          ))}
        </div>
      </div>
      <div className="list-row-side">
        <div className="row" style={{ gap: 6 }}>
          {wo.is_blocked && <BlockedBadge />}
          {wo.is_overdue && <OverdueBadge />}
          <PriorityBadge priority={wo.priority} />
          <StatusBadge status={wo.status} />
        </div>
        <AvatarStack users={wo.assignees} teams={wo.assignee_teams} />
      </div>
    </button>
  );
}

function WorkOrderTable({
  items,
  selectedId,
  onSelect,
}: {
  items: WorkOrderListItem[];
  selectedId: string | null;
  onSelect: (id: string) => void;
}) {
  return (
    <div className="table-wrap" data-testid="work-order-table">
      <table className="table">
        <thead>
          <tr>
            <th>#</th>
            <th>Title</th>
            <th>Status</th>
            <th>Priority</th>
            <th>Assignees</th>
            <th>Project</th>
            <th>Team</th>
            <th>Due</th>
            <th>Type</th>
          </tr>
        </thead>
        <tbody>
          {items.map((wo) => (
            <tr
              key={wo.id}
              className={cn('clickable', wo.id === selectedId && 'selected')}
              onClick={() => onSelect(wo.id)}
              tabIndex={0}
              onKeyDown={(e) => e.key === 'Enter' && onSelect(wo.id)}
            >
              <td className="num mono">{wo.number}</td>
              <td>
                <span className="row" style={{ gap: 6 }}>
                  <span style={{ fontWeight: 600 }}>{wo.title}</span>
                  {wo.is_blocked && <BlockedBadge />}
                </span>
              </td>
              <td>
                <StatusBadge status={wo.status} />
              </td>
              <td>
                <PriorityBadge priority={wo.priority} />
              </td>
              <td>
                <AvatarStack users={wo.assignees} teams={wo.assignee_teams} />
              </td>
              <td>{wo.project?.name || '—'}</td>
              <td>{wo.team?.name || '—'}</td>
              <td
                className={cn(wo.is_overdue && 'text-danger')}
                style={wo.is_overdue ? { color: 'var(--color-danger)', fontWeight: 600 } : undefined}
              >
                {fmtDue(wo.due_at)}
              </td>
              <td>
                <span className="row" style={{ gap: 4 }}>
                  {wo.work_type === 'PREVENTIVE' && <Repeat size={12} />}
                  {wo.work_type.toLowerCase().replace(/^./, (c) => c.toUpperCase())}
                </span>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}
