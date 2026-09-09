import { useId, useRef, type ReactNode } from 'react';
import { createPortal } from 'react-dom';
import { X } from 'lucide-react';
import { IconButton } from './Button';
import { useModalBehaviour } from './Dialog';
import { cn } from '@/lib/cn';

/**
 * Right-side pane used for create/edit flows so the list stays visible behind it.
 * Spec §7.2: 560px wide, header + scrolling body + sticky footer, Escape closes.
 */
export function SideSheet({
  open,
  onClose,
  title,
  subtitle,
  children,
  footer,
  wide = false,
  testId,
}: {
  open: boolean;
  onClose: () => void;
  title: ReactNode;
  subtitle?: ReactNode;
  children: ReactNode;
  footer?: ReactNode;
  wide?: boolean;
  testId?: string;
}) {
  const ref = useRef<HTMLDivElement>(null);
  const id = useId();
  useModalBehaviour(open, onClose, ref);
  if (!open) return null;
  return createPortal(
    <>
      <div className="sheet-overlay" onMouseDown={onClose} />
      <div
        ref={ref}
        className={cn('sheet', wide && 'sheet-wide')}
        role="dialog"
        aria-modal="true"
        aria-labelledby={`${id}-title`}
        data-testid={testId}
      >
        <div className="sheet-head">
          <div className="page-title-wrap">
            <h2 className="sheet-title" id={`${id}-title`}>
              {title}
            </h2>
            {subtitle && (
              <div className="text-muted text-label" style={{ fontWeight: 500 }}>
                {subtitle}
              </div>
            )}
          </div>
          <IconButton label="Close" onClick={onClose}>
            <X />
          </IconButton>
        </div>
        <div className="sheet-body">{children}</div>
        {footer && <div className="sheet-foot">{footer}</div>}
      </div>
    </>,
    document.body,
  );
}
