import { useCallback, useMemo } from 'react';
import { useSearchParams } from 'react-router-dom';

/**
 * URL-backed list state following the API conventions: `filter[key]=a,b`, `sort`, `q`,
 * plus page-specific extras. Keeping it in the URL makes views shareable and back-button safe.
 */
export interface ListState {
  filters: Record<string, string[]>;
  sort: string;
  q: string;
  extras: Record<string, string>;
}

export function useListParams(extraKeys: string[] = []) {
  const [params, setParams] = useSearchParams();

  const state = useMemo<ListState>(() => {
    const filters: Record<string, string[]> = {};
    const extras: Record<string, string> = {};
    params.forEach((value, key) => {
      if (key.startsWith('filter[') && key.endsWith(']')) {
        const values = value
          .split(',')
          .map((v) => v.trim())
          .filter(Boolean);
        if (values.length) filters[key.slice(7, -1)] = values;
      } else if (extraKeys.includes(key)) {
        extras[key] = value;
      }
    });
    return { filters, sort: params.get('sort') || '', q: params.get('q') || '', extras };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [params, extraKeys.join('|')]);

  const update = useCallback(
    (
      patch: {
        filters?: Record<string, string[] | undefined>;
        sort?: string | null;
        q?: string | null;
        extras?: Record<string, string | null | undefined>;
      },
      replace = true,
    ) => {
      const next = new URLSearchParams(params);
      if (patch.filters) {
        for (const [key, values] of Object.entries(patch.filters)) {
          if (!values || values.length === 0) next.delete(`filter[${key}]`);
          else next.set(`filter[${key}]`, values.join(','));
        }
      }
      if (patch.sort !== undefined) {
        if (patch.sort) next.set('sort', patch.sort);
        else next.delete('sort');
      }
      if (patch.q !== undefined) {
        if (patch.q) next.set('q', patch.q);
        else next.delete('q');
      }
      if (patch.extras) {
        for (const [key, value] of Object.entries(patch.extras)) {
          if (value === null || value === undefined || value === '') next.delete(key);
          else next.set(key, value);
        }
      }
      setParams(next, { replace });
    },
    [params, setParams],
  );

  const clearFilters = useCallback(() => {
    const next = new URLSearchParams(params);
    Array.from(next.keys()).forEach((key) => {
      if (key.startsWith('filter[') || key === 'q') next.delete(key);
    });
    setParams(next, { replace: true });
  }, [params, setParams]);

  /** Query-string parameters in the shape the API expects. */
  const apiParams = useMemo(() => {
    const out: Record<string, string> = {};
    for (const [key, values] of Object.entries(state.filters)) out[`filter[${key}]`] = values.join(',');
    if (state.sort) out.sort = state.sort;
    if (state.q) out.q = state.q;
    return out;
  }, [state]);

  const activeFilterCount = Object.keys(state.filters).length + (state.q ? 1 : 0);

  return { state, update, clearFilters, apiParams, activeFilterCount };
}
