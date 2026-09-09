import {
  cloneElement,
  useCallback,
  useEffect,
  useId,
  useRef,
  useState,
  type MouseEvent as ReactMouseEvent,
  type ReactElement,
  type ReactNode,
} from 'react';
import { Link } from 'react-router-dom';
import { cn } from '@/lib/cn';

/**
 * Minimal, dependency-free popover: the trigger toggles, outside click / Escape close,
 * focus returns to the trigger. Positioned relative to a wrapping inline-block.
 */
export function Popover({
  trigger,
  children,
  align = 'left',
  open: controlledOpen,
  onOpenChange,
  className,
  role = 'dialog',
  label,
}: {
  trigger: ReactElement;
  children: ReactNode | ((close: () => void) => ReactNode);
  align?: 'left' | 'right';
  open?: boolean;
  onOpenChange?: (open: boolean) => void;
  className?: string;
  role?: 'dialog' | 'menu' | 'listbox';
  label?: string;
}) {
  const [internalOpen, setInternalOpen] = useState(false);
  const open = controlledOpen ?? internalOpen;
  const setOpen = useCallback(
    (v: boolean) => {
      setInternalOpen(v);
      onOpenChange?.(v);
    },
    [onOpenChange],
  );
  const wrapRef = useRef<HTMLSpanElement>(null);
  const id = useId();

  useEffect(() => {
    if (!open) return;
    const onDown = (e: MouseEvent) => {
      if (wrapRef.current && !wrapRef.current.contains(e.target as Node)) setOpen(false);
    };
    const onKey = (e: KeyboardEvent) => {
      if (e.key === 'Escape') {
        e.stopPropagation();
        setOpen(false);
        (wrapRef.current?.querySelector('[aria-haspopup]') as HTMLElement | null)?.focus();
      }
    };
    document.addEventListener('mousedown', onDown);
    document.addEventListener('keydown', onKey);
    return () => {
      document.removeEventListener('mousedown', onDown);
      document.removeEventListener('keydown', onKey);
    };
  }, [open, setOpen]);

  const close = useCallback(() => setOpen(false), [setOpen]);

  return (
    <span ref={wrapRef} style={{ position: 'relative', display: 'inline-block', maxWidth: '100%' }}>
      {cloneElement(trigger, {
        onClick: (e: ReactMouseEvent) => {
          trigger.props.onClick?.(e);
          setOpen(!open);
        },
        'aria-haspopup': role === 'dialog' ? 'dialog' : role,
        'aria-expanded': open,
        'aria-controls': open ? id : undefined,
      })}
      {open && (
        <div
          id={id}
          role={role}
          aria-label={label}
          className={cn('popover', align === 'right' && 'popover-right', className)}
          style={{ top: 'calc(100% + 6px)' }}
        >
          {typeof children === 'function' ? children(close) : children}
        </div>
      )}
    </span>
  );
}

export interface MenuItemDef {
  key: string;
  label: ReactNode;
  icon?: ReactNode;
  onSelect?: () => void;
  to?: string;
  danger?: boolean;
  disabled?: boolean;
  separatorBefore?: boolean;
}

export function DropdownMenu({
  trigger,
  items,
  align = 'right',
  label,
}: {
  trigger: ReactElement;
  items: MenuItemDef[];
  align?: 'left' | 'right';
  label: string;
}) {
  return (
    <Popover trigger={trigger} align={align} role="menu" label={label}>
      {(close) => (
        <div className="menu">
          {items.map((item) => (
            <MenuRow key={item.key} item={item} close={close} />
          ))}
        </div>
      )}
    </Popover>
  );
}

function MenuRow({ item, close }: { item: MenuItemDef; close: () => void }) {
  const inner = (
    <>
      {item.icon}
      <span>{item.label}</span>
    </>
  );
  return (
    <>
      {item.separatorBefore && <div className="menu-separator" role="separator" />}
      {item.to ? (
        <Link
          className={cn('menu-item', item.danger && 'menu-item-danger')}
          to={item.to}
          role="menuitem"
          onClick={close}
        >
          {inner}
        </Link>
      ) : (
        <button
          type="button"
          className={cn('menu-item', item.danger && 'menu-item-danger')}
          role="menuitem"
          disabled={item.disabled}
          onClick={() => {
            close();
            item.onSelect?.();
          }}
        >
          {inner}
        </button>
      )}
    </>
  );
}

export function Tooltip({ text, children }: { text: string; children: ReactElement }) {
  const [show, setShow] = useState(false);
  return (
    <span
      style={{ position: 'relative', display: 'inline-flex' }}
      onMouseEnter={() => setShow(true)}
      onMouseLeave={() => setShow(false)}
      onFocus={() => setShow(true)}
      onBlur={() => setShow(false)}
    >
      {children}
      {show && (
        <span className="tooltip" role="tooltip" style={{ left: '50%', top: 0 }}>
          {text}
        </span>
      )}
    </span>
  );
}
