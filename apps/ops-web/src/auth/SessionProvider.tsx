import { createContext, useCallback, useContext, useEffect, useMemo, type ReactNode } from 'react';
import { Navigate, useLocation, useNavigate } from 'react-router-dom';
import { useQueryClient } from '@tanstack/react-query';
import { ApiError, api, onUnauthorized } from '@/api/client';
import { keys, useSession } from '@/api/hooks';
import type { Session } from '@/contracts/types';
import { ErrorState, Skeleton } from '@/ui';

interface SessionContextValue {
  session: Session;
  can: (permission: string) => boolean;
  canAny: (...permissions: string[]) => boolean;
  refresh: () => Promise<unknown>;
  logout: () => Promise<void>;
}

const SessionContext = createContext<SessionContextValue | null>(null);

export function useCurrentSession(): SessionContextValue {
  const ctx = useContext(SessionContext);
  if (!ctx) throw new Error('useCurrentSession must be used inside RequireAuth');
  return ctx;
}

/** Wraps authenticated routes: loads the session, redirects to /login on 401, exposes `can()`. */
export function RequireAuth({ children }: { children: ReactNode }) {
  const query = useSession();
  const location = useLocation();
  const navigate = useNavigate();
  const client = useQueryClient();

  useEffect(() => {
    return onUnauthorized(() => {
      client.clear();
      navigate(`/login?next=${encodeURIComponent(location.pathname + location.search)}`, { replace: true });
    });
  }, [client, navigate, location.pathname, location.search]);

  const logout = useCallback(async () => {
    try {
      await api('/ops/session/logout', { method: 'POST' });
    } finally {
      client.clear();
      navigate('/login', { replace: true });
    }
  }, [client, navigate]);

  const value = useMemo<SessionContextValue | null>(() => {
    if (!query.data) return null;
    const permissions = new Set(query.data.permissions);
    return {
      session: query.data,
      can: (permission) => permissions.has(permission),
      canAny: (...list) => list.some((p) => permissions.has(p)),
      refresh: () => client.invalidateQueries({ queryKey: keys.session }),
      logout,
    };
  }, [query.data, client, logout]);

  if (query.isPending) {
    return (
      <div style={{ padding: 40, maxWidth: 480 }}>
        <Skeleton lines={5} />
      </div>
    );
  }
  if (query.error instanceof ApiError && (query.error.status === 401 || query.error.status === 403)) {
    return (
      <Navigate
        to={`/login?next=${encodeURIComponent(location.pathname + location.search)}`}
        replace
        state={{ reason: query.error.code }}
      />
    );
  }
  if (query.error || !value) {
    return (
      <div style={{ padding: 40, maxWidth: 560 }}>
        <ErrorState error={query.error} onRetry={() => query.refetch()} />
      </div>
    );
  }
  return <SessionContext.Provider value={value}>{children}</SessionContext.Provider>;
}

/** Route guard for permission-gated pages; renders a friendly 403 rather than a blank page. */
export function RequirePermission({
  permission,
  children,
}: {
  permission: string | string[];
  children: ReactNode;
}) {
  const { canAny } = useCurrentSession();
  const list = Array.isArray(permission) ? permission : [permission];
  if (!canAny(...list)) {
    return (
      <div className="page">
        <div className="page-body" style={{ paddingTop: 32 }}>
          <ErrorState
            error={
              new ApiError(403, {
                ok: false,
                code: 'forbidden',
                error: 'Your role does not include this area. Ask a chapter admin if you need access.',
                permission: list[0],
              })
            }
          />
        </div>
      </div>
    );
  }
  return <>{children}</>;
}
