import { useEffect, useId, useRef, type ReactNode, type RefObject } from 'react';
import { createPortal } from 'react-dom';
import { X } from 'lucide-react';
import { IconButton, Button } from './Button';
import { cn } from '@/lib/cn';

const FOCUSABLE =
  'a[href],button:not([disabled]),input:not([disabled]),select:not([disabled]),textarea:not([disabled]),[tabindex]:not([tabindex="-1"])';

/** Shared modal behaviour: focus trap, Escape to close, scroll lock, focus restore. */
export function useModalBehaviour(open: boolean, onClose: () => void, ref: RefObject<HTMLElement>) {
  useEffect(() => {
    if (!open) return;
    const previous = document.activeElement as HTMLElement | null;
    const node = ref.current;
    const first =
      node?.querySelector<HTMLElement>('[data-autofocus]') || node?.querySelector<HTMLElement>(FOCUSABLE);
    first?.focus();
    const bodyOverflow = document.body.style.overflow;
    document.body.style.overflow = 'hidden';
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        // A popover (picker, filter chip, menu) open inside the modal owns this Escape.
        if (node?.querySelector('.popover')) return;
        e.stopPropagation();
        onClose();
        return;
      }
      if (e.key === 'Tab' && node) {
        const items = Array.from(node.querySelectorAll<HTMLElement>(FOCUSABLE)).filter(
          (el) => el.offsetParent !== null,
        );
        if (items.length === 0) return;
        const firstEl = items[0];
        const lastEl = items[items.length - 1];
        if (e.shiftKey && document.activeElement === firstEl) {
          e.preventDefault();
          lastEl.focus();
        } else if (!e.shiftKey && document.activeElement === lastEl) {
          e.preventDefault();
          firstEl.focus();
        }
      }
    };
    document.addEventListener('keydown', onKey);
    return () => {
      document.removeEventListener('keydown', onKey);
      document.body.style.overflow = bodyOverflow;
      previous?.focus?.();
    };
  }, [open, onClose, ref]);
}

export function Dialog({
  open,
  onClose,
  title,
  children,
  footer,
  size = 'md',
  description,
}: {
  open: boolean;
  onClose: () => void;
  title: ReactNode;
  children: ReactNode;
  footer?: ReactNode;
  size?: 'md' | 'lg';
  description?: ReactNode;
}) {
  const ref = useRef<HTMLDivElement>(null);
  const id = useId();
  useModalBehaviour(open, onClose, ref);
  if (!open) return null;
  return createPortal(
    <div className="overlay" onMouseDown={(e) => e.target === e.currentTarget && onClose()}>
      <div
        ref={ref}
        className={cn('dialog', size === 'lg' && 'dialog-lg')}
        role="dialog"
        aria-modal="true"
        aria-labelledby={`${id}-title`}
        aria-describedby={description ? `${id}-desc` : undefined}
      >
        <div className="dialog-head">
          <div>
            <h2 className="dialog-title" id={`${id}-title`}>
              {title}
            </h2>
            {description && (
              <p className="text-muted" id={`${id}-desc`} style={{ marginTop: 4 }}>
                {description}
              </p>
            )}
          </div>
          <IconButton label="Close" onClick={onClose}>
            <X />
          </IconButton>
        </div>
        <div className="dialog-body">{children}</div>
        {footer && <div className="dialog-foot">{footer}</div>}
      </div>
    </div>,
    document.body,
  );
}

export function ConfirmDialog({
  open,
  onClose,
  onConfirm,
  title,
  body,
  confirmLabel = 'Confirm',
  danger = false,
  loading = false,
}: {
  open: boolean;
  onClose: () => void;
  onConfirm: () => void;
  title: ReactNode;
  body: ReactNode;
  confirmLabel?: string;
  danger?: boolean;
  loading?: boolean;
}) {
  return (
    <Dialog
      open={open}
      onClose={onClose}
      title={title}
      footer={
        <>
          <Button variant="ghost" onClick={onClose} disabled={loading}>
            Cancel
          </Button>
          <Button
            variant={danger ? 'danger' : 'primary'}
            onClick={onConfirm}
            loading={loading}
            data-autofocus
          >
            {confirmLabel}
          </Button>
        </>
      }
    >
      <div>{body}</div>
    </Dialog>
  );
}
