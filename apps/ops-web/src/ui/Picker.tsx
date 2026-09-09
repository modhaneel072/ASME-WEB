import { useId, useMemo, useState, type ReactNode } from 'react';
import { Check, ChevronDown, X } from 'lucide-react';
import { Popover } from './Popover';
import { cn } from '@/lib/cn';

export interface PickerOption<V extends string | number = string> {
  value: V;
  label: string;
  hint?: string;
  color?: string | null;
  group?: string;
  disabled?: boolean;
  render?: ReactNode;
}

/**
 * Accessible multi/single select with search. Renders selected values as tags inside
 * the trigger and a searchable checkbox list in a popover.
 */
export function Picker<V extends string | number>({
  options,
  value,
  onChange,
  multiple = true,
  placeholder = 'Select…',
  label,
  error,
  required,
  hint,
  id,
  disabled,
  emptyText = 'No matches',
  allowClear = true,
  renderTag,
}: {
  options: PickerOption<V>[];
  value: V[];
  onChange: (next: V[]) => void;
  multiple?: boolean;
  placeholder?: string;
  label?: ReactNode;
  error?: string;
  required?: boolean;
  hint?: ReactNode;
  id?: string;
  disabled?: boolean;
  emptyText?: string;
  allowClear?: boolean;
  renderTag?: (option: PickerOption<V>) => ReactNode;
}) {
  const auto = useId();
  const fieldId = id || auto;
  const [query, setQuery] = useState('');
  const [open, setOpen] = useState(false);
  const selected = useMemo(() => options.filter((o) => value.includes(o.value)), [options, value]);
  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    return q
      ? options.filter((o) => o.label.toLowerCase().includes(q) || (o.hint || '').toLowerCase().includes(q))
      : options;
  }, [options, query]);
  const groups = useMemo(() => {
    const map = new Map<string, PickerOption<V>[]>();
    for (const o of filtered) {
      const key = o.group || '';
      if (!map.has(key)) map.set(key, []);
      map.get(key)!.push(o);
    }
    return Array.from(map.entries());
  }, [filtered]);

  const toggle = (v: V) => {
    if (multiple) onChange(value.includes(v) ? value.filter((x) => x !== v) : [...value, v]);
    else {
      onChange(value.includes(v) && allowClear ? [] : [v]);
      setOpen(false);
    }
  };

  return (
    <div className={cn('field', error && 'field-invalid')}>
      {label && (
        <label className="field-label" htmlFor={fieldId}>
          {label}
          {required && (
            <span className="req" aria-hidden="true">
              *
            </span>
          )}
        </label>
      )}
      <Popover
        open={open}
        onOpenChange={(v) => {
          setOpen(v);
          if (!v) setQuery('');
        }}
        role="dialog"
        label={typeof label === 'string' ? label : 'Options'}
        className="picker-popover"
        trigger={
          <button
            type="button"
            id={fieldId}
            className="picker-trigger"
            disabled={disabled}
            aria-invalid={!!error || undefined}
            aria-describedby={error ? `${fieldId}-error` : undefined}
            data-testid={`picker-${fieldId}`}
          >
            {selected.length === 0 ? (
              <span className="picker-placeholder">{placeholder}</span>
            ) : multiple ? (
              selected.map((o) => (
                <span key={String(o.value)} className="picker-tag">
                  {renderTag ? renderTag(o) : o.label}
                  <span
                    role="button"
                    tabIndex={0}
                    aria-label={`Remove ${o.label}`}
                    onClick={(e) => {
                      e.stopPropagation();
                      toggle(o.value);
                    }}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter' || e.key === ' ') {
                        e.preventDefault();
                        e.stopPropagation();
                        toggle(o.value);
                      }
                    }}
                    style={{ display: 'inline-flex' }}
                  >
                    <X size={12} />
                  </span>
                </span>
              ))
            ) : (
              <span className="truncate">{renderTag ? renderTag(selected[0]) : selected[0].label}</span>
            )}
            <span className="sr-only">
              <ChevronDown />
            </span>
          </button>
        }
      >
        <div style={{ width: 300, maxWidth: 'calc(100vw - 32px)' }}>
          {options.length > 6 && (
            <div className="picker-search">
              <input
                type="search"
                className="field-control"
                style={{ height: 32 }}
                placeholder="Type to filter"
                value={query}
                onChange={(e) => setQuery(e.target.value)}
                aria-label="Filter options"
                data-autofocus
                autoFocus
              />
            </div>
          )}
          <div className="picker-list" role="listbox" aria-multiselectable={multiple}>
            {filtered.length === 0 && <div className="picker-empty">{emptyText}</div>}
            {groups.map(([group, items]) => (
              <div key={group}>
                {group && <div className="menu-heading">{group}</div>}
                {items.map((o) => {
                  const isSelected = value.includes(o.value);
                  return (
                    <button
                      key={String(o.value)}
                      type="button"
                      role="option"
                      aria-selected={isSelected}
                      className="picker-option"
                      disabled={o.disabled}
                      onClick={() => toggle(o.value)}
                    >
                      <span className="picker-check" aria-hidden="true">
                        {isSelected && <Check size={16} />}
                      </span>
                      {o.color && (
                        <span className="chip-dot" style={{ background: o.color }} aria-hidden="true" />
                      )}
                      <span style={{ flex: 1, minWidth: 0 }}>
                        <span className="truncate" style={{ display: 'block' }}>
                          {o.render || o.label}
                        </span>
                        {o.hint && <span className="text-caption text-muted">{o.hint}</span>}
                      </span>
                    </button>
                  );
                })}
              </div>
            ))}
          </div>
          {multiple && value.length > 0 && (
            <div style={{ padding: 8, borderTop: '1px solid var(--color-border)' }}>
              <button type="button" className="link-button" onClick={() => onChange([])}>
                Clear selection
              </button>
            </div>
          )}
        </div>
      </Popover>
      {error ? (
        <div className="field-error" id={`${fieldId}-error`} role="alert">
          {error}
        </div>
      ) : hint ? (
        <div className="field-hint">{hint}</div>
      ) : null}
    </div>
  );
}
