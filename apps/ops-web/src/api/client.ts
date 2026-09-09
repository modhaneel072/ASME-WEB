import type { ApiErrorBody } from '@/contracts/types';

export class ApiError extends Error {
  code: string;
  status: number;
  details: { field: string | null; message: string }[];
  permission?: string;
  body: ApiErrorBody | null;

  constructor(status: number, body: ApiErrorBody | null, fallback = 'Request failed') {
    super(body?.error || fallback);
    this.name = 'ApiError';
    this.status = status;
    this.code = body?.code || (status === 0 ? 'network' : 'http_error');
    this.details = body?.details || [];
    this.permission = body?.permission;
    this.body = body;
  }

  /** Field -> message map for react-hook-form setError. */
  fieldErrors(): Record<string, string> {
    const out: Record<string, string> = {};
    for (const d of this.details) if (d.field && !out[d.field]) out[d.field] = d.message;
    return out;
  }
}

export const CSRF_HEADER: Record<string, string> = { 'X-Requested-With': 'ASME-Ops' };

type Listener = (error: ApiError) => void;
const unauthorizedListeners = new Set<Listener>();

export function onUnauthorized(listener: Listener) {
  unauthorizedListeners.add(listener);
  return () => {
    unauthorizedListeners.delete(listener);
  };
}

interface RequestOptions {
  method?: 'GET' | 'POST' | 'PATCH' | 'PUT' | 'DELETE';
  body?: unknown;
  form?: FormData;
  signal?: AbortSignal;
}

/** Calls the ops API and unwraps the `{ ok, payload }` envelope. */
export async function api<T>(path: string, options: RequestOptions = {}): Promise<T> {
  const { method = 'GET', body, form, signal } = options;
  const headers: Record<string, string> = { Accept: 'application/json', ...CSRF_HEADER };
  let payload: BodyInit | undefined;
  if (form) payload = form;
  else if (body !== undefined) {
    headers['Content-Type'] = 'application/json';
    payload = JSON.stringify(body);
  }
  let response: Response;
  try {
    response = await fetch(`/api/v1${path}`, {
      method,
      headers,
      body: payload,
      credentials: 'same-origin',
      signal,
    });
  } catch (err) {
    if ((err as Error).name === 'AbortError') throw err;
    throw new ApiError(0, null, 'You appear to be offline. Check your connection and retry.');
  }
  let json: any = null;
  const text = await response.text();
  if (text) {
    try {
      json = JSON.parse(text);
    } catch {
      json = null;
    }
  }
  if (!response.ok || (json && json.ok === false)) {
    const error = new ApiError(response.status, json);
    if (response.status === 401) unauthorizedListeners.forEach((l) => l(error));
    throw error;
  }
  return (json && 'payload' in json ? json.payload : json) as T;
}

export function qs(params: Record<string, string | number | boolean | string[] | null | undefined>): string {
  const search = new URLSearchParams();
  for (const [key, value] of Object.entries(params)) {
    if (value === undefined || value === null || value === '' || (Array.isArray(value) && value.length === 0))
      continue;
    search.set(key, Array.isArray(value) ? value.join(',') : String(value));
  }
  const s = search.toString();
  return s ? `?${s}` : '';
}
