import type { ReactNode } from 'react';
import { AlertTriangle, Inbox, WifiOff } from 'lucide-react';
import { Button } from './Button';
import { ApiError } from '@/api/client';

export function EmptyState({
  icon,
  title,
  body,
  actions,
}: {
  icon?: ReactNode;
  title: ReactNode;
  body?: ReactNode;
  actions?: ReactNode;
}) {
  return (
    <div className="empty" data-testid="empty-state">
      <div className="empty-icon">{icon || <Inbox />}</div>
      <div className="empty-title">{title}</div>
      {body && <div className="empty-body">{body}</div>}
      {actions && <div className="empty-actions">{actions}</div>}
    </div>
  );
}

export function Skeleton({
  lines = 4,
  widths = ['70%', '45%', '85%', '60%'],
}: {
  lines?: number;
  widths?: string[];
}) {
  return (
    <div className="skeleton" aria-busy="true" aria-live="polite" aria-label="Loading">
      {Array.from({ length: lines }).map((_, i) => (
        <div key={i} className="skeleton-line" style={{ width: widths[i % widths.length] }} />
      ))}
    </div>
  );
}

export function ErrorState({
  error,
  onRetry,
  title,
}: {
  error: unknown;
  onRetry?: () => void;
  title?: string;
}) {
  const apiError = error instanceof ApiError ? error : null;
  const offline = apiError?.status === 0;
  const forbidden = apiError?.status === 403;
  const message = apiError?.message || (error instanceof Error ? error.message : 'Something went wrong.');
  return (
    <div className="error-state" role="alert">
      <div className="row" style={{ fontWeight: 650 }}>
        {offline ? <WifiOff size={18} /> : <AlertTriangle size={18} />}
        {title ||
          (offline
            ? "You're offline"
            : forbidden
              ? "You don't have access to this"
              : 'Could not load this view')}
      </div>
      <div>{message}</div>
      {apiError?.permission && <div className="text-caption">Requires permission: {apiError.permission}</div>}
      {onRetry && !forbidden && (
        <div>
          <Button size="sm" onClick={onRetry}>
            Retry
          </Button>
        </div>
      )}
    </div>
  );
}
