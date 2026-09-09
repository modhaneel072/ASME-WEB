import { afterEach, describe, expect, it, vi } from 'vitest';
import { api, ApiError, onUnauthorized, qs } from './client';

function mockFetch(status: number, body: unknown) {
  const fn = vi.fn(async () => ({ ok: status < 400, status, text: async () => JSON.stringify(body) }));
  vi.stubGlobal('fetch', fn);
  return fn;
}

afterEach(() => vi.unstubAllGlobals());

describe('api client', () => {
  it('unwraps the ok/payload envelope and sends the CSRF header', async () => {
    const fetchMock = mockFetch(200, { ok: true, payload: { id: 'x' } });
    const result = await api<{ id: string }>('/work-orders/x', { method: 'PATCH', body: { title: 'T' } });
    expect(result).toEqual({ id: 'x' });
    const [url, init] = fetchMock.mock.calls[0] as unknown as [string, RequestInit];
    expect(url).toBe('/api/v1/work-orders/x');
    expect((init.headers as Record<string, string>)['X-Requested-With']).toBe('ASME-Ops');
    expect((init.headers as Record<string, string>)['Content-Type']).toBe('application/json');
    expect(init.credentials).toBe('same-origin');
  });

  it('throws a typed ApiError with field details on validation failures', async () => {
    mockFetch(400, {
      ok: false,
      code: 'validation',
      error: 'Check the highlighted fields.',
      details: [
        { field: 'title', message: 'Title is required' },
        { field: null, message: 'ignored' },
      ],
    });
    const err = (await api('/work-orders', { method: 'POST', body: {} }).catch((e) => e)) as ApiError;
    expect(err).toBeInstanceOf(ApiError);
    expect(err.code).toBe('validation');
    expect(err.status).toBe(400);
    expect(err.fieldErrors()).toEqual({ title: 'Title is required' });
  });

  it('notifies unauthorized listeners on 401', async () => {
    mockFetch(401, { ok: false, code: 'login_required', error: 'Login required.' });
    const listener = vi.fn();
    const off = onUnauthorized(listener);
    await expect(api('/ops/session')).rejects.toBeInstanceOf(ApiError);
    expect(listener).toHaveBeenCalledTimes(1);
    off();
  });

  it('maps network failures to a friendly offline error', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(async () => {
        throw new TypeError('Failed to fetch');
      }),
    );
    const err = (await api('/ops/session').catch((e) => e)) as ApiError;
    expect(err).toBeInstanceOf(ApiError);
    expect(err.status).toBe(0);
    expect(err.code).toBe('network');
  });

  it('builds query strings following the list conventions', () => {
    expect(
      qs({
        'filter[status]': ['OPEN', 'IN_PROGRESS'],
        sort: 'due_asc',
        q: '',
        empty: undefined,
        'page[cursor]': null,
      }),
    ).toBe('?filter%5Bstatus%5D=OPEN%2CIN_PROGRESS&sort=due_asc');
    expect(qs({})).toBe('');
  });
});
