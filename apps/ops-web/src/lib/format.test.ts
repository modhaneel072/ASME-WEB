import { describe, expect, it } from 'vitest';
import { fmtBytes, fmtMinutes, humanize, parse, toLocalInput } from './format';

describe('format helpers', () => {
  it('formats minutes as hours and minutes', () => {
    expect(fmtMinutes(0)).toBe('0m');
    expect(fmtMinutes(45)).toBe('45m');
    expect(fmtMinutes(60)).toBe('1h');
    expect(fmtMinutes(135)).toBe('2h 15m');
  });

  it('humanizes enum keys', () => {
    expect(humanize('IN_PROGRESS')).toBe('In Progress');
    expect(humanize('on_hold')).toBe('On Hold');
    expect(humanize(null)).toBe('');
  });

  it('parses naive server timestamps as UTC', () => {
    const d = parse('2026-09-10T14:00:00');
    expect(d?.toISOString()).toBe('2026-09-10T14:00:00.000Z');
    expect(parse('2026-09-10T14:00:00Z')?.toISOString()).toBe('2026-09-10T14:00:00.000Z');
    expect(parse('2026-09-10')).not.toBeNull();
    expect(parse(null)).toBeNull();
    expect(parse('not a date')).toBeNull();
  });

  it('round-trips into datetime-local inputs', () => {
    const local = toLocalInput('2026-09-10T14:00:00');
    expect(local).toMatch(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}$/);
    expect(toLocalInput(null)).toBe('');
  });

  it('formats byte sizes', () => {
    expect(fmtBytes(512)).toBe('512 B');
    expect(fmtBytes(2048)).toBe('2.0 KB');
    expect(fmtBytes(3 * 1024 * 1024)).toBe('3.0 MB');
  });
});
