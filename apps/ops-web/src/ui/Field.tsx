import {
  forwardRef,
  useId,
  type InputHTMLAttributes,
  type ReactNode,
  type SelectHTMLAttributes,
  type TextareaHTMLAttributes,
} from 'react';
import { cn } from '@/lib/cn';

interface FieldWrapProps {
  label?: ReactNode;
  required?: boolean;
  error?: string;
  hint?: ReactNode;
  className?: string;
  id: string;
  children: ReactNode;
}

export function FieldWrap({ label, required, error, hint, className, id, children }: FieldWrapProps) {
  return (
    <div className={cn('field', error && 'field-invalid', className)}>
      {label && (
        <label className="field-label" htmlFor={id}>
          {label}
          {required && (
            <span className="req" aria-hidden="true">
              *
            </span>
          )}
        </label>
      )}
      {children}
      {error ? (
        <div className="field-error" id={`${id}-error`} role="alert">
          {error}
        </div>
      ) : hint ? (
        <div className="field-hint" id={`${id}-hint`}>
          {hint}
        </div>
      ) : null}
    </div>
  );
}

type Common = {
  label?: ReactNode;
  required?: boolean;
  error?: string;
  hint?: ReactNode;
  wrapClassName?: string;
};

export const TextField = forwardRef<HTMLInputElement, InputHTMLAttributes<HTMLInputElement> & Common>(
  function TextField({ label, required, error, hint, wrapClassName, id, className, ...rest }, ref) {
    const auto = useId();
    const fieldId = id || auto;
    return (
      <FieldWrap
        label={label}
        required={required}
        error={error}
        hint={hint}
        className={wrapClassName}
        id={fieldId}
      >
        <input
          ref={ref}
          id={fieldId}
          className={cn('field-control', className)}
          aria-invalid={!!error || undefined}
          aria-describedby={error ? `${fieldId}-error` : hint ? `${fieldId}-hint` : undefined}
          aria-required={required || undefined}
          {...rest}
        />
      </FieldWrap>
    );
  },
);

export const TextArea = forwardRef<HTMLTextAreaElement, TextareaHTMLAttributes<HTMLTextAreaElement> & Common>(
  function TextArea({ label, required, error, hint, wrapClassName, id, className, ...rest }, ref) {
    const auto = useId();
    const fieldId = id || auto;
    return (
      <FieldWrap
        label={label}
        required={required}
        error={error}
        hint={hint}
        className={wrapClassName}
        id={fieldId}
      >
        <textarea
          ref={ref}
          id={fieldId}
          className={cn('field-control', className)}
          aria-invalid={!!error || undefined}
          aria-describedby={error ? `${fieldId}-error` : undefined}
          {...rest}
        />
      </FieldWrap>
    );
  },
);

export interface SelectOption {
  value: string;
  label: string;
  disabled?: boolean;
}

export const SelectField = forwardRef<
  HTMLSelectElement,
  SelectHTMLAttributes<HTMLSelectElement> & Common & { options: SelectOption[]; placeholder?: string }
>(function SelectField(
  { label, required, error, hint, wrapClassName, id, className, options, placeholder, ...rest },
  ref,
) {
  const auto = useId();
  const fieldId = id || auto;
  return (
    <FieldWrap
      label={label}
      required={required}
      error={error}
      hint={hint}
      className={wrapClassName}
      id={fieldId}
    >
      <select
        ref={ref}
        id={fieldId}
        className={cn('field-control', className)}
        aria-invalid={!!error || undefined}
        aria-describedby={error ? `${fieldId}-error` : undefined}
        {...rest}
      >
        {placeholder !== undefined && <option value="">{placeholder}</option>}
        {options.map((o) => (
          <option key={o.value} value={o.value} disabled={o.disabled}>
            {o.label}
          </option>
        ))}
      </select>
    </FieldWrap>
  );
});

export const Checkbox = forwardRef<
  HTMLInputElement,
  InputHTMLAttributes<HTMLInputElement> & { label: ReactNode }
>(function Checkbox({ label, className, ...rest }, ref) {
  return (
    <label className={cn('checkbox', className)}>
      <input ref={ref} type="checkbox" {...rest} />
      <span>{label}</span>
    </label>
  );
});

export function Segmented<T extends string>({
  value,
  onChange,
  options,
  label,
}: {
  value: T;
  onChange: (v: T) => void;
  options: { value: T; label: ReactNode }[];
  label: string;
}) {
  return (
    <div className="segmented" role="group" aria-label={label}>
      {options.map((o) => (
        <button
          key={o.value}
          type="button"
          aria-pressed={o.value === value}
          onClick={() => onChange(o.value)}
        >
          {o.label}
        </button>
      ))}
    </div>
  );
}
