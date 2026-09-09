import { describe, expect, it } from 'vitest';
import { completeSchema, projectSchema, workOrderCreateSchema } from './schemas';

describe('work order create schema', () => {
  it('normalises empty optional fields to undefined so PATCH/POST bodies stay clean', () => {
    const out = workOrderCreateSchema.parse({
      title: '  Inspect hub  ',
      description: '',
      project_id: '',
      estimated_minutes: Number.NaN,
      due_at: '',
      recurrence: null,
    });
    expect(out.title).toBe('Inspect hub');
    expect(out.description).toBeUndefined();
    expect(out.project_id).toBeUndefined();
    expect(out.estimated_minutes).toBeUndefined();
    expect(out.due_at).toBeUndefined();
    expect(out.recurrence).toBeUndefined();
    expect(out.status).toBe('OPEN');
    expect(out.priority).toBe('NONE');
    expect(out.assignee_user_ids).toEqual([]);
  });

  it('converts datetime-local values to ISO-8601 UTC', () => {
    const out = workOrderCreateSchema.parse({ title: 'x', due_at: '2026-09-10T14:00' });
    expect(out.due_at).toMatch(/Z$/);
    expect(new Date(out.due_at!).getTime()).toBe(new Date('2026-09-10T14:00').getTime());
  });

  it('rejects a due date before the start date and an empty title', () => {
    const bad = workOrderCreateSchema.safeParse({
      title: 'x',
      start_at: '2026-09-12T10:00',
      due_at: '2026-09-10T10:00',
    });
    expect(bad.success).toBe(false);
    if (!bad.success) expect(bad.error.issues[0].path).toEqual(['due_at']);
    const empty = workOrderCreateSchema.safeParse({ title: '   ' });
    expect(empty.success).toBe(false);
  });
});

describe('complete schema', () => {
  it('keeps costs and drops blank asset status', () => {
    const out = completeSchema.parse({
      completion_note: 'done',
      time_minutes: 30,
      costs: [{ type: 'part', amount: 12.5, description: '' }],
      asset_status: '',
    });
    expect(out.asset_status).toBeUndefined();
    expect(out.costs[0]).toEqual({ type: 'part', amount: 12.5, description: undefined });
  });
});

describe('project schema', () => {
  it('requires a name and defaults status/risk', () => {
    const out = projectSchema.parse({
      name: 'Crater Cruncher Rover',
      budget_amount: Number.NaN,
      lead_user_id: null,
    });
    expect(out.status).toBe('active');
    expect(out.risk_level).toBe('medium');
    expect(out.budget_amount).toBeUndefined();
    expect(out.lead_user_id).toBeUndefined();
    expect(projectSchema.safeParse({ name: 'x' }).success).toBe(false);
  });
});
