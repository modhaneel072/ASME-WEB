import { act, renderHook } from '@testing-library/react';
import { MemoryRouter } from 'react-router-dom';
import { describe, expect, it } from 'vitest';
import type { ReactNode } from 'react';
import { useListParams } from './listParams';

const wrapper = (initial: string) =>
  function Wrapper({ children }: { children: ReactNode }) {
    return <MemoryRouter initialEntries={[initial]}>{children}</MemoryRouter>;
  };

describe('useListParams', () => {
  it('reads filter[], sort, q and extras from the URL', () => {
    const { result } = renderHook(() => useListParams(['tab']), {
      wrapper: wrapper(
        '/work-orders?filter[status]=OPEN,IN_PROGRESS&filter[assignee]=me&sort=priority_desc&q=hub&tab=done',
      ),
    });
    expect(result.current.state.filters).toEqual({ status: ['OPEN', 'IN_PROGRESS'], assignee: ['me'] });
    expect(result.current.state.sort).toBe('priority_desc');
    expect(result.current.state.q).toBe('hub');
    expect(result.current.state.extras.tab).toBe('done');
    expect(result.current.apiParams).toEqual({
      'filter[status]': 'OPEN,IN_PROGRESS',
      'filter[assignee]': 'me',
      sort: 'priority_desc',
      q: 'hub',
    });
    expect(result.current.activeFilterCount).toBe(3);
  });

  it('updates and clears filters without touching extras', () => {
    const { result } = renderHook(() => useListParams(['tab']), {
      wrapper: wrapper('/work-orders?tab=done&filter[status]=OPEN'),
    });
    act(() => result.current.update({ filters: { priority: ['HIGH'] }, sort: 'due_asc' }));
    expect(result.current.state.filters).toEqual({ status: ['OPEN'], priority: ['HIGH'] });
    expect(result.current.state.sort).toBe('due_asc');
    act(() => result.current.update({ filters: { status: [] } }));
    expect(result.current.state.filters).toEqual({ priority: ['HIGH'] });
    act(() => result.current.clearFilters());
    expect(result.current.state.filters).toEqual({});
    expect(result.current.state.extras.tab).toBe('done');
    expect(result.current.state.sort).toBe('due_asc');
  });
});
