import { useEffect, useState, type ReactNode } from 'react';
import { NavLink, Outlet, useLocation } from 'react-router-dom';
import {
  Bell,
  BarChart3,
  Boxes,
  ClipboardList,
  FolderKanban,
  LogOut,
  MapPin,
  Menu,
  Rocket,
  Settings,
  Tags,
  Users,
  ExternalLink,
  Globe,
  X,
  Check,
} from 'lucide-react';
import { useCurrentSession } from '@/auth/SessionProvider';
import { api } from '@/api/client';
import { keys, useApiMutation, useNotifications, useSetup } from '@/api/hooks';
import { Avatar, DropdownMenu, IconButton, Popover, Button } from '@/ui';
import { fmtRelative } from '@/lib/format';
import { cn } from '@/lib/cn';
import type { Notification } from '@/contracts/types';

interface NavItem {
  to: string;
  label: string;
  icon: ReactNode;
  permission?: string | string[];
  testId: string;
}

interface NavGroup {
  title: string;
  items: NavItem[];
}

// Only modules that exist in this release are listed (spec §5: no decorative navigation).
const GROUPS: NavGroup[] = [
  {
    title: 'Setup',
    items: [
      {
        to: '/setup',
        label: 'Setup Center',
        icon: <Rocket />,
        permission: 'setup.manage',
        testId: 'nav-setup',
      },
    ],
  },
  {
    title: 'Work',
    items: [
      {
        to: '/work-orders',
        label: 'Work Orders',
        icon: <ClipboardList />,
        permission: 'work_order.read_assigned',
        testId: 'nav-work-orders',
      },
      {
        to: '/projects',
        label: 'Projects',
        icon: <FolderKanban />,
        permission: 'project.read',
        testId: 'nav-projects',
      },
    ],
  },
  {
    title: 'Optimize',
    items: [
      {
        to: '/reporting',
        label: 'Reporting',
        icon: <BarChart3 />,
        permission: 'report.view',
        testId: 'nav-reporting',
      },
    ],
  },
  {
    title: 'Manage',
    items: [
      { to: '/assets', label: 'Assets', icon: <Boxes />, permission: 'asset.read', testId: 'nav-assets' },
      { to: '/locations', label: 'Locations', icon: <MapPin />, testId: 'nav-locations' },
      { to: '/categories', label: 'Categories', icon: <Tags />, testId: 'nav-categories' },
      {
        to: '/teams-users',
        label: 'Teams & Users',
        icon: <Users />,
        permission: 'team.read',
        testId: 'nav-teams-users',
      },
      {
        to: '/settings',
        label: 'Settings',
        icon: <Settings />,
        permission: ['org.read', 'audit.read', 'org.manage'],
        testId: 'nav-settings',
      },
    ],
  },
];

export function AppShell() {
  const { session, canAny, logout } = useCurrentSession();
  const [drawerOpen, setDrawerOpen] = useState(false);
  const location = useLocation();

  useEffect(() => setDrawerOpen(false), [location.pathname]);
  useEffect(() => {
    if (!drawerOpen) return;
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') setDrawerOpen(false);
    };
    document.addEventListener('keydown', onKey);
    return () => document.removeEventListener('keydown', onKey);
  }, [drawerOpen]);

  const visibleGroups = GROUPS.map((g) => ({
    ...g,
    items: g.items.filter(
      (i) => !i.permission || canAny(...(Array.isArray(i.permission) ? i.permission : [i.permission])),
    ),
  })).filter((g) => g.items.length > 0);

  return (
    <div className="shell">
      {drawerOpen && <div className="drawer-overlay" onClick={() => setDrawerOpen(false)} />}
      <aside
        className={cn('sidebar', drawerOpen && 'open')}
        aria-label="Primary navigation"
        data-testid="sidebar"
      >
        <div className="sidebar-brand">
          <div className="sidebar-brand-mark" aria-hidden="true">
            UI
          </div>
          <div className="sidebar-brand-text">
            <div className="sidebar-brand-name">ASME Ops</div>
            <div className="sidebar-brand-org truncate">{session.organization.name}</div>
          </div>
          <IconButton
            label="Close navigation"
            className="topbar-only"
            onClick={() => setDrawerOpen(false)}
            style={{ marginLeft: 'auto' }}
          >
            <X />
          </IconButton>
        </div>
        <nav className="sidebar-nav">
          {visibleGroups.map((group) => (
            <div key={group.title}>
              <div className="sidebar-group-title">{group.title}</div>
              {group.items.map((item) => (
                <NavLink
                  key={item.to}
                  to={item.to}
                  className={({ isActive }) => cn('sidebar-link', isActive && 'active')}
                  data-testid={item.testId}
                  onClick={() => setDrawerOpen(false)}
                >
                  {item.icon}
                  <span>{item.label}</span>
                  {item.to === '/setup' && <SetupNavBadge />}
                </NavLink>
              ))}
            </div>
          ))}
        </nav>
        <div className="sidebar-foot">
          <div className="row" style={{ gap: 4 }}>
            <div style={{ flex: 1, minWidth: 0 }}>
              <DropdownMenu
                label="Account menu"
                align="left"
                trigger={
                  <button type="button" className="sidebar-user" data-testid="user-menu">
                    <Avatar user={session.user} />
                    <span style={{ minWidth: 0 }}>
                      <span className="sidebar-user-name truncate" style={{ display: 'block' }}>
                        {session.user.name}
                      </span>
                      <span className="sidebar-user-role">{session.role.name}</span>
                    </span>
                  </button>
                }
                items={[
                  {
                    key: 'legacy',
                    label: 'Legacy portal',
                    icon: <ExternalLink />,
                    onSelect: () => window.location.assign('/legacy/app'),
                  },
                  {
                    key: 'public',
                    label: 'Public site',
                    icon: <Globe />,
                    onSelect: () => window.location.assign('/'),
                  },
                  {
                    key: 'logout',
                    label: 'Sign out',
                    icon: <LogOut />,
                    onSelect: () => void logout(),
                    separatorBefore: true,
                  },
                ]}
              />
            </div>
            <NotificationsMenu />
          </div>
        </div>
      </aside>
      <div className="main">
        <header className="topbar">
          <IconButton label="Open navigation" onClick={() => setDrawerOpen(true)} data-testid="open-nav">
            <Menu />
          </IconButton>
          <span className="topbar-title">ASME Ops</span>
        </header>
        <SetupBanner />
        <main className="content" id="main">
          <Outlet />
        </main>
      </div>
    </div>
  );
}

function SetupNavBadge() {
  const { data } = useSetup();
  if (!data || data.complete) return null;
  return <span className="sidebar-badge">{data.steps_left}</span>;
}

/** Spec §7.1: persistent banner until setup is complete or the user dismisses it. */
function SetupBanner() {
  const { session, can, refresh } = useCurrentSession();
  const { data } = useSetup();
  const dismiss = useApiMutation(
    () => api('/ops/setup/banner', { method: 'POST', body: { dismissed: true } }),
    [keys.setup],
  );
  if (
    !can('setup.manage') ||
    !data ||
    data.complete ||
    session.setup_banner_dismissed ||
    data.banner_dismissed
  )
    return null;
  return (
    <div className="setup-banner" role="region" aria-label="Setup progress" data-testid="setup-banner">
      <strong>Setup {data.percent}% complete</strong>
      <div className="progress gold" aria-hidden="true">
        <span style={{ width: `${data.percent}%` }} />
      </div>
      <span>
        {data.steps_left} step{data.steps_left === 1 ? '' : 's'} left
        {data.next_step && (
          <>
            {' · '}
            <NavLink to={data.next_step.route.replace(/^\/app/, '')}>Next: {data.next_step.title}</NavLink>
          </>
        )}
      </span>
      <button
        type="button"
        className="link-button dismiss"
        onClick={async () => {
          await dismiss.mutateAsync(undefined as never);
          await refresh();
        }}
      >
        Hide
      </button>
    </div>
  );
}

function NotificationsMenu() {
  const { data } = useNotifications();
  const readAll = useApiMutation(
    () => api('/notifications/read-all', { method: 'POST' }),
    [keys.notifications, keys.session],
  );
  const readOne = useApiMutation(
    (id: string) => api(`/notifications/${id}/read`, { method: 'POST' }),
    [keys.notifications, keys.session],
  );
  const unread = data?.unread ?? 0;
  return (
    <Popover
      align="right"
      label="Notifications"
      trigger={
        <button
          type="button"
          className="btn btn-ghost btn-icon"
          aria-label={unread ? `Notifications, ${unread} unread` : 'Notifications'}
          data-testid="notifications"
          style={{ position: 'relative' }}
        >
          <Bell />
          {unread > 0 && (
            <span
              className="sidebar-badge"
              style={{ position: 'absolute', top: 2, right: 2, minWidth: 16, height: 16, fontSize: 10 }}
            >
              {unread}
            </span>
          )}
        </button>
      }
    >
      <div style={{ width: 340, maxWidth: 'calc(100vw - 32px)' }}>
        <div className="card-head">
          <span className="card-title" style={{ fontSize: 14 }}>
            Notifications
          </span>
          {unread > 0 && (
            <Button
              size="sm"
              variant="ghost"
              icon={<Check />}
              onClick={() => readAll.mutate(undefined as never)}
            >
              Mark all read
            </Button>
          )}
        </div>
        <div className="picker-list" style={{ maxHeight: 360 }}>
          {(data?.items || []).length === 0 && (
            <div className="picker-empty">Nothing yet. Assignments and status changes will appear here.</div>
          )}
          {(data?.items || []).map((n: Notification) => (
            <NavLink
              key={n.id}
              to={
                n.entity_type === 'work_order' && n.entity_id
                  ? `/work-orders?wo=${n.entity_id}`
                  : n.entity_type === 'project' && n.entity_id
                    ? `/projects/${n.entity_id}`
                    : '#'
              }
              className="picker-option"
              style={{
                alignItems: 'flex-start',
                background: n.read_at ? undefined : 'var(--color-bg-selected)',
              }}
              onClick={() => !n.read_at && readOne.mutate(n.id)}
            >
              <span style={{ minWidth: 0 }}>
                <span className="text-label" style={{ display: 'block' }}>
                  {n.title}
                </span>
                {n.body && <span className="text-caption text-muted">{n.body}</span>}
                <span className="text-caption text-muted" style={{ display: 'block' }}>
                  {fmtRelative(n.created_at)}
                </span>
              </span>
            </NavLink>
          ))}
        </div>
      </div>
    </Popover>
  );
}
