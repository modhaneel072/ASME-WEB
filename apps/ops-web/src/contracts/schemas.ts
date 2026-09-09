import { z } from 'zod';

// Zod schemas for forms. They mirror the pydantic request models; the server remains
// the authority and its field errors are merged back into the form.

export const loginSchema = z.object({
  identifier: z.string().trim().min(1, 'Enter your email or username').max(160),
  password: z.string().min(1, 'Enter your password').max(200),
});
export type LoginInput = z.infer<typeof loginSchema>;

const optionalStr = (max: number) =>
  z
    .string()
    .trim()
    .max(max)
    .optional()
    .or(z.literal(''))
    .transform((v) => (v ? v : undefined));
const optionalId = z
  .string()
  .optional()
  .or(z.literal(''))
  .transform((v) => (v ? v : undefined));
const optionalInt = z
  .union([z.number().int().min(0), z.nan(), z.literal('')])
  .optional()
  .transform((v) => (typeof v === 'number' && !Number.isNaN(v) ? v : undefined));
const optionalDateTime = z
  .string()
  .optional()
  .or(z.literal(''))
  .transform((v) => (v ? new Date(v).toISOString() : undefined));
const optionalDate = z
  .string()
  .optional()
  .or(z.literal(''))
  .transform((v) => (v ? v : undefined));

export const WO_PRIORITIES = ['NONE', 'LOW', 'MEDIUM', 'HIGH', 'CRITICAL'] as const;
export const WO_WORK_TYPES = [
  'REACTIVE',
  'PREVENTIVE',
  'PROJECT',
  'EVENT',
  'INSPECTION',
  'SAFETY',
  'PROCUREMENT',
  'DOCUMENTATION',
] as const;

export const workOrderCreateSchema = z
  .object({
    title: z.string().trim().min(1, 'Title is required').max(220),
    description: optionalStr(20000),
    project_id: optionalId,
    location_id: optionalId,
    primary_asset_id: optionalId,
    asset_ids: z.array(z.string()).default([]),
    assignee_user_ids: z.array(z.number()).default([]),
    assignee_team_ids: z.array(z.string()).default([]),
    team_id: optionalId,
    estimated_minutes: optionalInt,
    due_at: optionalDateTime,
    start_at: optionalDateTime,
    recurrence: z
      .object({
        frequency: z.enum(['daily', 'weekly', 'monthly']),
        interval: z.number().int().min(1).max(52),
        mode: z.enum(['fixed', 'floating']),
      })
      .optional()
      .nullable()
      .transform((v) => v ?? undefined),
    work_type: z.enum(WO_WORK_TYPES).default('PROJECT'),
    priority: z.enum(WO_PRIORITIES).default('NONE'),
    category_ids: z.array(z.string()).default([]),
    budget_code: optionalStr(60),
    watcher_user_ids: z.array(z.number()).default([]),
    parent_work_order_id: optionalId,
    status: z.enum(['DRAFT', 'OPEN']).default('OPEN'),
    parent_completion_policy: z.enum(['manual', 'auto']).default('manual'),
  })
  .refine((v) => !(v.due_at && v.start_at) || new Date(v.due_at) >= new Date(v.start_at), {
    message: 'Due date cannot be before the start date',
    path: ['due_at'],
  });
export type WorkOrderCreateInput = z.input<typeof workOrderCreateSchema>;
export type WorkOrderCreateOutput = z.output<typeof workOrderCreateSchema>;

export const subWorkOrderSchema = z.object({
  title: z.string().trim().min(1, 'Title is required').max(220),
  description: optionalStr(20000),
  assignee_user_ids: z.array(z.number()).default([]),
  due_at: optionalDateTime,
  priority: z.enum(WO_PRIORITIES).default('NONE'),
  estimated_minutes: optionalInt,
});

export const completeSchema = z.object({
  completion_note: optionalStr(5000),
  time_minutes: optionalInt,
  costs: z
    .array(
      z.object({
        type: z.enum(['part', 'labor', 'vendor', 'other']),
        amount: z.number().min(0),
        description: optionalStr(300),
      }),
    )
    .default([]),
  asset_status: z
    .enum(['ONLINE', 'OFFLINE_PLANNED', 'OFFLINE_UNPLANNED', 'DO_NOT_TRACK', 'RETIRED'])
    .optional()
    .or(z.literal(''))
    .transform((v) => (v ? v : undefined)),
  follow_up_title: optionalStr(220),
});
export type CompleteInput = z.input<typeof completeSchema>;

export const projectSchema = z.object({
  name: z.string().trim().min(2, 'Name must be at least 2 characters').max(200),
  code: optionalStr(30),
  description: optionalStr(10000),
  status: z.enum(['planning', 'active', 'on_hold', 'completed', 'archived']).default('active'),
  risk_level: z.enum(['low', 'medium', 'high']).default('medium'),
  lead_user_id: z
    .number()
    .optional()
    .nullable()
    .transform((v) => v ?? undefined),
  faculty_advisor_user_id: z
    .number()
    .optional()
    .nullable()
    .transform((v) => v ?? undefined),
  competition: optionalStr(200),
  academic_year: optionalStr(9),
  start_date: optionalDate,
  target_date: optionalDate,
  budget_amount: z
    .union([z.number().min(0), z.nan(), z.literal('')])
    .optional()
    .transform((v) => (typeof v === 'number' && !Number.isNaN(v) ? v : undefined)),
  budget_code: optionalStr(60),
  repository_url: optionalStr(500),
  cad_url: optionalStr(500),
  requirements_url: optionalStr(500),
  visibility: z.enum(['private', 'members', 'public']).default('members'),
});
export type ProjectInput = z.input<typeof projectSchema>;

export const milestoneSchema = z.object({
  name: z.string().trim().min(1, 'Name is required').max(200),
  description: optionalStr(5000),
  due_date: optionalDate,
  status: z.enum(['planned', 'in_progress', 'complete', 'missed']).default('planned'),
  owner_user_id: z
    .number()
    .optional()
    .nullable()
    .transform((v) => v ?? undefined),
  weight: z.number().int().min(1).max(100).default(1),
});
export type MilestoneInput = z.input<typeof milestoneSchema>;

export const teamSchema = z.object({
  name: z.string().trim().min(2, 'Name must be at least 2 characters').max(160),
  description: optionalStr(2000),
  parent_team_id: optionalId,
  project_id: optionalId,
  color: optionalStr(20),
  escalation_note: optionalStr(500),
});
export type TeamInput = z.input<typeof teamSchema>;

export const locationSchema = z.object({
  name: z.string().trim().min(1, 'Name is required').max(160),
  description: optionalStr(2000),
  parent_location_id: optionalId,
  building: optionalStr(160),
  room: optionalStr(80),
});
export type LocationInput = z.input<typeof locationSchema>;

export const categorySchema = z.object({
  name: z.string().trim().min(1, 'Name is required').max(120),
  color: z
    .string()
    .regex(/^#[0-9a-fA-F]{6}$/, 'Use a hex colour like #0878d1')
    .default('#0878d1'),
  icon: z.string().trim().min(1).max(60).default('tag'),
  description: optionalStr(2000),
});
export type CategoryInput = z.input<typeof categorySchema>;

export const assetSchema = z.object({
  name: z.string().trim().min(1, 'Name is required').max(200),
  code: optionalStr(60),
  description: optionalStr(5000),
  parent_asset_id: optionalId,
  project_id: optionalId,
  location_id: optionalId,
  responsible_team_id: optionalId,
  owner_user_id: z
    .number()
    .optional()
    .nullable()
    .transform((v) => v ?? undefined),
  manufacturer: optionalStr(160),
  model: optionalStr(160),
  serial_number: optionalStr(160),
  purchase_date: optionalDate,
  purchase_cost: z
    .union([z.number().min(0), z.nan(), z.literal('')])
    .optional()
    .transform((v) => (typeof v === 'number' && !Number.isNaN(v) ? v : undefined)),
  warranty_end: optionalDate,
  criticality: z.enum(['low', 'medium', 'high', 'critical']).default('medium'),
  status: z
    .enum(['ONLINE', 'OFFLINE_PLANNED', 'OFFLINE_UNPLANNED', 'DO_NOT_TRACK', 'RETIRED'])
    .default('ONLINE'),
  asset_type_ids: z.array(z.string()).default([]),
});
export type AssetInput = z.input<typeof assetSchema>;

export const inviteSchema = z.object({
  name: z.string().trim().min(2).max(160),
  email: z.string().trim().email('Enter a valid email'),
  role_key: z.string().min(2),
  password: z
    .string()
    .min(8, 'At least 8 characters')
    .max(200)
    .optional()
    .or(z.literal(''))
    .transform((v) => (v ? v : undefined)),
  title: optionalStr(120),
});
export type InviteInput = z.input<typeof inviteSchema>;

export const savedFilterSchema = z.object({
  name: z.string().trim().min(1, 'Give the view a name').max(120),
  visibility: z.enum(['private', 'team', 'chapter']).default('private'),
  team_id: optionalId,
  is_default: z.boolean().default(false),
});
export type SavedFilterInput = z.input<typeof savedFilterSchema>;
