import { createContext, useCallback, useContext, useMemo, useRef, useState, type ReactNode } from 'react';
import { AlertCircle, CheckCircle2, Info } from 'lucide-react';
import { cn } from '@/lib/cn';

interface ToastItem {
  id: number;
  tone: 'success' | 'error' | 'info';
  title: string;
  message?: string;
  action?: { label: string; onClick: () => void };
}

interface ToastApi {
  success: (title: string, message?: string, action?: ToastItem['action']) => void;
  error: (title: string, message?: string) => void;
  info: (title: string, message?: string) => void;
}

const ToastContext = createContext<ToastApi | null>(null);

export function ToastProvider({ children }: { children: ReactNode }) {
  const [items, setItems] = useState<ToastItem[]>([]);
  const counter = useRef(0);

  const push = useCallback((item: Omit<ToastItem, 'id'>) => {
    const id = ++counter.current;
    setItems((prev) => [...prev.slice(-3), { ...item, id }]);
    window.setTimeout(
      () => setItems((prev) => prev.filter((t) => t.id !== id)),
      item.tone === 'error' ? 7000 : 4500,
    );
  }, []);

  const api = useMemo<ToastApi>(
    () => ({
      success: (title, message, action) => push({ tone: 'success', title, message, action }),
      error: (title, message) => push({ tone: 'error', title, message }),
      info: (title, message) => push({ tone: 'info', title, message }),
    }),
    [push],
  );

  return (
    <ToastContext.Provider value={api}>
      {children}
      <div className="toast-region" aria-live="polite" aria-atomic="false">
        {items.map((t) => (
          <div
            key={t.id}
            className={cn('toast', `toast-${t.tone}`)}
            role={t.tone === 'error' ? 'alert' : 'status'}
          >
            {t.tone === 'success' ? <CheckCircle2 /> : t.tone === 'error' ? <AlertCircle /> : <Info />}
            <div className="toast-body">
              <div className="toast-title">{t.title}</div>
              {t.message && <div className="toast-message">{t.message}</div>}
            </div>
            {t.action && (
              <button type="button" className="toast-action" onClick={t.action.onClick}>
                {t.action.label}
              </button>
            )}
          </div>
        ))}
      </div>
    </ToastContext.Provider>
  );
}

export function useToast(): ToastApi {
  const ctx = useContext(ToastContext);
  if (!ctx) throw new Error('useToast must be used inside ToastProvider');
  return ctx;
}
