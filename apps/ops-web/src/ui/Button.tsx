import { forwardRef, type ButtonHTMLAttributes, type ReactNode } from 'react';
import { Link } from 'react-router-dom';
import { ChevronDown } from 'lucide-react';
import { cn } from '@/lib/cn';

type Variant = 'primary' | 'secondary' | 'ghost' | 'danger' | 'danger-outline' | 'success';

export interface ButtonProps extends ButtonHTMLAttributes<HTMLButtonElement> {
  variant?: Variant;
  size?: 'md' | 'sm';
  loading?: boolean;
  icon?: ReactNode;
  block?: boolean;
}

export const Button = forwardRef<HTMLButtonElement, ButtonProps>(function Button(
  {
    variant = 'secondary',
    size = 'md',
    loading = false,
    icon,
    block,
    className,
    children,
    disabled,
    type = 'button',
    ...rest
  },
  ref,
) {
  return (
    <button
      ref={ref}
      type={type}
      className={cn('btn', `btn-${variant}`, size === 'sm' && 'btn-sm', block && 'btn-block', className)}
      disabled={disabled || loading}
      aria-busy={loading || undefined}
      {...rest}
    >
      {loading ? <span className="spinner" aria-hidden="true" /> : icon}
      {children}
    </button>
  );
});

export interface IconButtonProps extends ButtonHTMLAttributes<HTMLButtonElement> {
  label: string;
  variant?: Variant;
  size?: 'md' | 'sm';
}

export const IconButton = forwardRef<HTMLButtonElement, IconButtonProps>(function IconButton(
  { label, variant = 'ghost', size = 'md', className, children, type = 'button', ...rest },
  ref,
) {
  return (
    <button
      ref={ref}
      type={type}
      className={cn('btn btn-icon', `btn-${variant}`, size === 'sm' && 'btn-sm', className)}
      aria-label={label}
      title={label}
      {...rest}
    >
      {children}
    </button>
  );
});

export function LinkButton({
  to,
  variant = 'secondary',
  size = 'md',
  icon,
  className,
  children,
}: {
  to: string;
  variant?: Variant;
  size?: 'md' | 'sm';
  icon?: ReactNode;
  className?: string;
  children: ReactNode;
}) {
  return (
    <Link to={to} className={cn('btn', `btn-${variant}`, size === 'sm' && 'btn-sm', className)}>
      {icon}
      {children}
    </Link>
  );
}

/** Primary action with a secondary chevron that opens a menu of related actions. */
export function SplitButton({
  children,
  onClick,
  menuLabel,
  onMenu,
  loading,
  disabled,
  variant = 'primary',
  icon,
}: {
  children: ReactNode;
  onClick: () => void;
  menuLabel: string;
  onMenu: () => void;
  loading?: boolean;
  disabled?: boolean;
  variant?: Variant;
  icon?: ReactNode;
}) {
  return (
    <span className="btn-split">
      <Button variant={variant} onClick={onClick} loading={loading} disabled={disabled} icon={icon}>
        {children}
      </Button>
      <button
        type="button"
        className={cn('btn', `btn-${variant}`)}
        aria-label={menuLabel}
        onClick={onMenu}
        disabled={disabled}
      >
        <ChevronDown />
      </button>
    </span>
  );
}
