// Mirrors the pydantic contracts in asme/ops/schemas and the serializers in asme/ops/services.
// Source of truth: GET /api/v1/ops/openapi.json.

export interface UserRef {
  id: number;
  name: string;
  email: string;
  initials: string;
  role_key?: string | null;
  role_name?: string | null;
}

export interface Ref {
  id: string;
  name: string;
}

export interface Organization {
  id: string;
  name: string;
  slug: string;
  timezone: string;
  academic_year_start_month: number;
  logo_url: string | null;
  settings: Record<string, unknown>;
}

export interface Session {
  user: UserRef;
  organization: Organization;
  role: { key: string; name: string; is_custom: boolean };
  permissions: string[];
  grants: Record<string, string>;
  setup_banner_dismissed: boolean;
  unread_notifications: number;
  title?: string | null;
}

export interface Lookups {
  users: UserRef[];
  teams: {
    id: string;
    name: string;
    color: string | null;
    project_id: string | null;
    lead_ids: number[];
    member_ids: number[];
  }[];
  locations: { id: string; name: string; parent_location_id: string | null; is_default: boolean }[];
  categories: { id: string; name: string; color: string; icon: string }[];
  projects: { id: string; name: string; code: string; status: string; lead_user_id: number | null }[];
  assets: {
    id: string;
    name: string;
    code: string | null;
    parent_asset_id: string | null;
    project_id: string | null;
    location_id: string | null;
    status: string;
  }[];
  roles: { key: string; name: string }[];
}

export type WoStatus = 'DRAFT' | 'OPEN' | 'IN_PROGRESS' | 'ON_HOLD' | 'DONE' | 'CANCELED' | 'SKIPPED';
export type WoPriority = 'NONE' | 'LOW' | 'MEDIUM' | 'HIGH' | 'CRITICAL';
export type WoWorkType =
  'REACTIVE' | 'PREVENTIVE' | 'PROJECT' | 'EVENT' | 'INSPECTION' | 'SAFETY' | 'PROCUREMENT' | 'DOCUMENTATION';

export interface CategoryChip {
  id: string;
  name: string;
  color: string;
  icon: string;
}

export interface WorkOrderListItem {
  id: string;
  number: number;
  title: string;
  status: WoStatus;
  priority: WoPriority;
  work_type: WoWorkType;
  is_blocked: boolean;
  is_overdue: boolean;
  due_at: string | null;
  start_at: string | null;
  project: Ref | null;
  team: Ref | null;
  location: Ref | null;
  primary_asset: Ref | null;
  assignees: UserRef[];
  assignee_teams: Ref[];
  categories: CategoryChip[];
  parent_work_order_id: string | null;
  child_count: number;
  comment_count: number;
  estimated_minutes: number | null;
  updated_at: string;
  last_activity_at: string;
  created_at: string;
}

export interface StatusHistoryEntry {
  id: string;
  from_status: WoStatus | null;
  to_status: WoStatus;
  changed_by: UserRef | null;
  note: string | null;
  changed_at: string;
}

export interface TimeEntry {
  id: string;
  user: UserRef;
  minutes: number;
  started_at: string | null;
  ended_at: string | null;
  note: string | null;
  created_at: string;
}

export interface CostEntry {
  id: string;
  type: 'part' | 'labor' | 'vendor' | 'other';
  amount: number;
  description: string | null;
  created_at: string;
}

export interface Comment {
  id: string;
  entity_type: string;
  entity_id: string;
  body: string;
  deleted: boolean;
  author: UserRef;
  parent_comment_id: string | null;
  created_at: string;
  edited_at: string | null;
}

export interface Attachment {
  id: string;
  entity_type: string;
  entity_id: string;
  original_name: string;
  content_type: string;
  size_bytes: number;
  kind: 'file' | 'image';
  scan_status: string;
  uploaded_by: UserRef | null;
  created_at: string;
  download_url: string;
  expires_in_seconds: number;
}

export interface WorkOrderPermissions {
  edit: boolean;
  assign: boolean;
  start: boolean;
  complete: boolean;
  cancel: boolean;
  comment: boolean;
  upload: boolean;
  log_time: boolean;
  log_cost: boolean;
  set_critical: boolean;
}

export interface WorkOrderDetail extends WorkOrderListItem {
  description: string | null;
  actual_minutes: number;
  completed_at: string | null;
  canceled_at: string | null;
  completion_note: string | null;
  cancel_reason: string | null;
  budget_code: string | null;
  recurrence: {
    frequency: 'daily' | 'weekly' | 'monthly';
    interval: number;
    mode: 'fixed' | 'floating';
  } | null;
  parent_completion_policy: 'manual' | 'auto';
  parent: { id: string; number: number; title: string } | null;
  creator: UserRef | null;
  watchers: UserRef[];
  related_assets: Ref[];
  children: WorkOrderListItem[];
  child_progress: {
    total: number;
    done: number;
    percent: number;
    estimated_minutes: number;
    actual_minutes: number;
    by_status: Record<string, number>;
  };
  status_history: StatusHistoryEntry[];
  time_entries: TimeEntry[];
  cost_entries: CostEntry[];
  total_cost: number;
  allowed_transitions: WoStatus[];
  permissions: WorkOrderPermissions;
  comments?: Comment[];
  files?: Attachment[];
  follow_up?: WorkOrderListItem | null;
  next_occurrence?: WorkOrderListItem | null;
}

export interface ListResponse<T> {
  items: T[];
  next_cursor: string | null;
  counts?: { todo: number; done: number };
}

export interface Milestone {
  id: string;
  project_id: string;
  name: string;
  description: string | null;
  due_date: string | null;
  status: 'planned' | 'in_progress' | 'complete' | 'missed';
  owner: UserRef | null;
  weight: number;
  order: number;
  completed_at: string | null;
  is_overdue: boolean;
}

export interface ProjectListItem {
  id: string;
  name: string;
  code: string;
  status: 'planning' | 'active' | 'on_hold' | 'completed' | 'archived';
  risk_level: 'low' | 'medium' | 'high';
  lead: UserRef | null;
  completion_percent: number;
  open_work_orders: number;
  overdue_work_orders: number;
  blocked_work_orders: number;
  next_milestone: { id: string; name: string; due_date: string; status: string; days_left: number } | null;
  target_date: string | null;
  academic_year: string | null;
  competition: string | null;
  team_count: number;
  visibility: string;
  archived_at: string | null;
  updated_at: string;
}

export interface ProjectMember extends UserRef {
  project_role: 'lead' | 'member' | 'advisor' | 'observer';
  team_id: string | null;
  team: Ref | null;
}

export interface ProjectDetail extends ProjectListItem {
  description: string | null;
  faculty_advisor: UserRef | null;
  start_date: string | null;
  budget_amount: number | null;
  budget_code: string | null;
  budget_used: number;
  hours_logged: number;
  repository_url: string | null;
  cad_url: string | null;
  requirements_url: string | null;
  public_project_id: number | null;
  public_slug: string | null;
  members: ProjectMember[];
  milestones: Milestone[];
  teams: (Ref & { member_count: number; color: string | null })[];
  created_at: string;
  permissions?: { edit: boolean; create_work_order: boolean };
  files?: Attachment[];
}

export interface ProjectHealth {
  project_id: string;
  completion_percent: number;
  work_orders: {
    open: number;
    done: number;
    total: number;
    overdue: number;
    blocked: number;
    by_status: Record<string, number>;
    by_priority: Record<string, number>;
  };
  workload_by_team: { team_id: string; team: string; open_work_orders: number }[];
  milestones: {
    total: number;
    complete: number;
    at_risk: Milestone[];
    next: ProjectListItem['next_milestone'];
  };
  budget: { amount: number | null; used: number; remaining: number | null };
  hours_logged: number;
  risk_level: string;
  generated_at: string;
}

export interface TeamMember extends UserRef {
  is_lead: boolean;
}

export interface Team {
  id: string;
  name: string;
  description: string | null;
  color: string | null;
  escalation_note: string | null;
  parent_team_id: string | null;
  parent: Ref | null;
  project_id: string | null;
  project: Ref | null;
  members: TeamMember[];
  lead_ids: number[];
  member_count: number;
  archived_at: string | null;
  created_at: string;
  updated_at: string;
}

export interface Member extends UserRef {
  membership_id: string;
  member_status: 'active' | 'invited' | 'suspended';
  title: string | null;
  joined_at: string | null;
  last_login_at: string | null;
  is_active: boolean;
}

export interface Location {
  id: string;
  name: string;
  description: string | null;
  parent_location_id: string | null;
  parent: Ref | null;
  building: string | null;
  room: string | null;
  is_default: boolean;
  qr_code: string | null;
  asset_count: number;
  archived_at: string | null;
  created_at: string;
  updated_at: string;
}

export interface Category extends CategoryChip {
  description: string | null;
  work_order_count: number;
  archived_at: string | null;
  created_at: string;
  updated_at: string;
}

export type AssetStatus = 'ONLINE' | 'OFFLINE_PLANNED' | 'OFFLINE_UNPLANNED' | 'DO_NOT_TRACK' | 'RETIRED';

export interface Asset {
  id: string;
  name: string;
  code: string | null;
  description: string | null;
  parent_asset_id: string | null;
  parent: Ref | null;
  project_id: string | null;
  project: Ref | null;
  location_id: string | null;
  location: Ref | null;
  responsible_team_id: string | null;
  responsible_team: Ref | null;
  owner: UserRef | null;
  manufacturer: string | null;
  model: string | null;
  serial_number: string | null;
  purchase_date: string | null;
  purchase_cost: number | null;
  warranty_end: string | null;
  criticality: 'low' | 'medium' | 'high' | 'critical';
  status: AssetStatus;
  qr_code: string | null;
  asset_types: { id: string; name: string; color: string; icon: string }[];
  custom_fields: Record<string, unknown>;
  child_count: number;
  open_work_orders: number;
  archived_at: string | null;
  created_at: string;
  updated_at: string;
  children?: Asset[];
  files?: Attachment[];
  permissions?: { edit: boolean };
}

export interface AssetStatusHistoryEntry {
  id: string;
  from_status: AssetStatus | null;
  to_status: AssetStatus;
  downtime_type: string | null;
  downtime_reason: string | null;
  note: string | null;
  started_at: string;
  ended_at: string | null;
  changed_by: UserRef | null;
}

export interface SetupTask {
  key: string;
  title: string;
  description: string;
  minutes: number;
  required: boolean;
  route: string;
  available: boolean;
  current: number;
  target: number;
  complete: boolean;
}

export interface SetupPhase {
  key: string;
  title: string;
  description: string;
  tasks: SetupTask[];
  required_total: number;
  required_done: number;
  complete: boolean;
  estimated_minutes: number;
  available_tasks: number;
}

export interface SetupState {
  phases: SetupPhase[];
  percent: number;
  required_total: number;
  required_done: number;
  steps_left: number;
  next_step: { phase: string; task: string; title: string; route: string } | null;
  banner_dismissed: boolean;
  complete: boolean;
}

export interface DashboardData {
  range: { key: string; start: string; end: string; days: number };
  project_id: string | null;
  totals: {
    open: number;
    overdue: number;
    blocked: number;
    due_soon: number;
    created: number;
    completed: number;
  };
  by_status: Record<string, number>;
  by_priority: Record<string, number>;
  by_work_type: Record<string, number>;
  repeating: { repeating: number; non_repeating: number };
  created_vs_completed: { start: string; created: number; completed: number }[];
  on_time_completion_rate: number | null;
  average_completion_hours: number | null;
  workload_by_team: { team_id: string; team: string; open: number }[];
  workload_by_user: { user_id: number; name: string; open: number }[];
  hours_logged: number;
  costs: { parts: number; labor: number; vendor: number; other: number; total: number };
  generated_at: string;
}

export interface Notification {
  id: string;
  type: string;
  title: string;
  body: string | null;
  entity_type: string | null;
  entity_id: string | null;
  read_at: string | null;
  created_at: string;
}

export interface SavedFilter {
  id: string;
  entity_type: string;
  name: string;
  visibility: 'private' | 'team' | 'chapter';
  team_id: string | null;
  filters: Record<string, string[]>;
  sort: string | null;
  view_type: 'panel' | 'table';
  is_default: boolean;
  owner: UserRef | null;
  is_mine: boolean;
  created_at: string;
}

export interface AuditEvent {
  id: string;
  event_type: string;
  entity_type: string;
  entity_id: string;
  actor: UserRef | null;
  before: Record<string, unknown> | null;
  after: Record<string, unknown> | null;
  metadata: Record<string, any>;
  occurred_at: string;
}

export interface Role {
  id: string;
  key: string | null;
  name: string;
  rank: number;
  is_custom: boolean;
  permissions: [string, string][];
}

export interface ApiErrorBody {
  ok: false;
  code: string;
  error: string;
  details?: { field: string | null; message: string; type?: string }[];
  permission?: string;
  [key: string]: unknown;
}
