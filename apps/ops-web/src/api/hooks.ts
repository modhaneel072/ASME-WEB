import {
  useInfiniteQuery,
  useMutation,
  useQuery,
  useQueryClient,
  type InfiniteData,
  type UseQueryOptions,
} from '@tanstack/react-query';
import { api, qs } from './client';
import type {
  Asset,
  AssetStatusHistoryEntry,
  AuditEvent,
  Category,
  DashboardData,
  ListResponse,
  Location,
  Lookups,
  Member,
  Notification,
  ProjectDetail,
  ProjectHealth,
  ProjectListItem,
  Role,
  SavedFilter,
  Session,
  SetupState,
  Team,
  WorkOrderDetail,
  WorkOrderListItem,
  Attachment,
} from '@/contracts/types';

export const keys = {
  session: ['session'] as const,
  lookups: ['lookups'] as const,
  setup: ['setup'] as const,
  notifications: ['notifications'] as const,
  workOrders: (params: Record<string, unknown>) => ['work-orders', params] as const,
  workOrder: (id: string) => ['work-order', id] as const,
  projects: (params: Record<string, unknown>) => ['projects', params] as const,
  project: (id: string) => ['project', id] as const,
  projectHealth: (id: string) => ['project-health', id] as const,
  projectActivity: (id: string) => ['project-activity', id] as const,
  teams: ['teams'] as const,
  users: (inactive: boolean) => ['users', inactive] as const,
  roles: ['roles'] as const,
  locations: ['locations'] as const,
  categories: ['categories'] as const,
  assets: (params: Record<string, unknown>) => ['assets', params] as const,
  asset: (id: string) => ['asset', id] as const,
  assetHistory: (id: string) => ['asset-history', id] as const,
  dashboard: (range: string, projectId: string | null) => ['dashboard', range, projectId] as const,
  savedFilters: (type: string) => ['saved-filters', type] as const,
  audit: (params: Record<string, unknown>) => ['audit', params] as const,
  files: (type: string, id: string) => ['files', type, id] as const,
};

export function useSession(options?: Partial<UseQueryOptions<Session>>) {
  return useQuery<Session>({
    queryKey: keys.session,
    queryFn: () => api<Session>('/ops/session'),
    retry: false,
    staleTime: 60_000,
    ...options,
  });
}

export function useLookups() {
  return useQuery<Lookups>({
    queryKey: keys.lookups,
    queryFn: () => api<Lookups>('/ops/lookups'),
    staleTime: 30_000,
  });
}

export function useSetup() {
  return useQuery<SetupState>({ queryKey: keys.setup, queryFn: () => api<SetupState>('/ops/setup') });
}

export function useNotifications() {
  return useQuery<{ items: Notification[]; unread: number }>({
    queryKey: keys.notifications,
    queryFn: () => api('/notifications'),
    refetchInterval: 30_000,
  });
}

export function useWorkOrders(params: Record<string, string | string[] | undefined>) {
  return useQuery<ListResponse<WorkOrderListItem>>({
    queryKey: keys.workOrders(params),
    queryFn: () => api(`/work-orders${qs(params)}`),
    placeholderData: (prev) => prev,
  });
}

export function useWorkOrder(id: string | null) {
  return useQuery<WorkOrderDetail>({
    queryKey: keys.workOrder(id || ''),
    queryFn: () => api(`/work-orders/${id}`),
    enabled: !!id,
  });
}

export function useProjects(params: Record<string, string | string[] | undefined>) {
  return useQuery<ListResponse<ProjectListItem>>({
    queryKey: keys.projects(params),
    queryFn: () => api(`/projects${qs(params)}`),
    placeholderData: (prev) => prev,
  });
}

export function useProject(id: string | null) {
  return useQuery<ProjectDetail>({
    queryKey: keys.project(id || ''),
    queryFn: () => api(`/projects/${id}`),
    enabled: !!id,
  });
}

export function useProjectHealth(id: string | null) {
  return useQuery<ProjectHealth>({
    queryKey: keys.projectHealth(id || ''),
    queryFn: () => api(`/projects/${id}/health`),
    enabled: !!id,
  });
}

export function useProjectActivity(id: string | null) {
  return useQuery<AuditEvent[]>({
    queryKey: keys.projectActivity(id || ''),
    queryFn: () => api(`/projects/${id}/activity`),
    enabled: !!id,
  });
}

export function useTeams() {
  return useQuery<Team[]>({ queryKey: keys.teams, queryFn: () => api('/teams') });
}

export function useUsers(includeInactive = false) {
  return useQuery<Member[]>({
    queryKey: keys.users(includeInactive),
    queryFn: () => api(`/users${includeInactive ? '?include_inactive=1' : ''}`),
  });
}

export function useRoles() {
  return useQuery<Role[]>({ queryKey: keys.roles, queryFn: () => api('/roles'), staleTime: 300_000 });
}

export function useLocations() {
  return useQuery<Location[]>({ queryKey: keys.locations, queryFn: () => api('/locations') });
}

export function useCategories() {
  return useQuery<Category[]>({ queryKey: keys.categories, queryFn: () => api('/categories') });
}

export function useAssets(params: Record<string, string | string[] | undefined>) {
  return useQuery<ListResponse<Asset>>({
    queryKey: keys.assets(params),
    queryFn: () => api(`/assets${qs(params)}`),
    placeholderData: (prev) => prev,
  });
}

export function useAsset(id: string | null) {
  return useQuery<Asset>({
    queryKey: keys.asset(id || ''),
    queryFn: () => api(`/assets/${id}`),
    enabled: !!id,
  });
}

export function useAssetHistory(id: string | null) {
  return useQuery<AssetStatusHistoryEntry[]>({
    queryKey: keys.assetHistory(id || ''),
    queryFn: () => api(`/assets/${id}/history`),
    enabled: !!id,
  });
}

export function useDashboard(range: string, projectId: string | null) {
  return useQuery<DashboardData>({
    queryKey: keys.dashboard(range, projectId),
    queryFn: () => api(`/ops/dashboard/operations${qs({ range, project_id: projectId })}`),
  });
}

export function useSavedFilters(entityType: string) {
  return useQuery<{ mine: SavedFilter[]; shared: SavedFilter[] }>({
    queryKey: keys.savedFilters(entityType),
    queryFn: () => api(`/saved-filters?entity_type=${entityType}`),
  });
}

export function useAudit(params: Record<string, string | undefined>) {
  return useQuery<{ items: AuditEvent[]; next_offset: number | null }>({
    queryKey: keys.audit(params),
    queryFn: () => api(`/ops/audit${qs(params)}`),
    placeholderData: (prev) => prev,
  });
}

export function useFiles(entityType: string, entityId: string | null) {
  return useQuery<Attachment[]>({
    queryKey: keys.files(entityType, entityId || ''),
    queryFn: () => api(`/files?entity_type=${entityType}&entity_id=${entityId}`),
    enabled: !!entityId,
  });
}

/** Generic mutation that invalidates the given query keys on success. */
export function useApiMutation<TVars, TResult = unknown>(
  fn: (vars: TVars) => Promise<TResult>,
  invalidate: (readonly unknown[])[] | ((vars: TVars, result: TResult) => (readonly unknown[])[]) = [],
) {
  const client = useQueryClient();
  return useMutation<TResult, Error, TVars>({
    mutationFn: fn,
    onSuccess: (result, vars) => {
      const list = typeof invalidate === 'function' ? invalidate(vars, result) : invalidate;
      for (const key of list) client.invalidateQueries({ queryKey: key });
    },
  });
}

export function useInvalidate() {
  const client = useQueryClient();
  return (...keysToInvalidate: (readonly unknown[])[]) =>
    keysToInvalidate.forEach((k) => client.invalidateQueries({ queryKey: k }));
}

/** Cursor-paginated work-order list (spec §9: `page[cursor]`). */
export function useWorkOrdersInfinite(params: Record<string, string | undefined>) {
  return useInfiniteQuery<
    ListResponse<WorkOrderListItem>,
    Error,
    InfiniteData<ListResponse<WorkOrderListItem>>,
    readonly unknown[],
    string | null
  >({
    queryKey: keys.workOrders(params),
    queryFn: ({ pageParam }) =>
      api(`/work-orders${qs({ ...params, 'page[cursor]': pageParam || undefined })}`),
    initialPageParam: null,
    getNextPageParam: (last) => last.next_cursor,
    placeholderData: (prev) => prev,
  });
}
