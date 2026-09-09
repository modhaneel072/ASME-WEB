import { useMemo, useState, type ReactNode } from 'react';
import { Check, ChevronDown, X } from 'lucide-react';
import { Popover } from './Popover';
import { cn } from '@/lib/cn';

export interface FilterOption {
  value: string;
  label: string;
  color?: string | null;
}

/**
 * A filter "chip" (spec §7.4): dashed when inactive, solid blue when values are selected.
 * Opens a searchable checklist; multi-select unless `single`.
 */
export function FilterChip({
  label,
  options,
  value,
  onChange,
  single = false,
  icon,
}: {
  label: string;
  options: FilterOption[];
  value: string[];
  onChange: (values: string[]) => void;
  single?: boolean;
  icon?: ReactNode;
}) {
  const [query, setQuery] = useState('');
  const active = value.length > 0;
  const selectedLabels = options.filter((o) => value.includes(o.value)).map((o) => o.label);
  const text = !active
    ? label
    : selectedLabels.length === 1
      ? `${label}: ${selectedLabels[0]}`
      : `${label}: ${selectedLabels.length}`;
  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    return q ? options.filter((o) => o.label.toLowerCase().includes(q)) : options;
  }, [options, query]);

  const toggle = (v: string, close: () => void) => {
    if (single) {
      onChange(value.includes(v) ? [] : [v]);
      close();
    } else onChange(value.includes(v) ? value.filter((x) => x !== v) : [...value, v]);
  };

  return (
    <Popover
      role="dialog"
      label={`${label} filter`}
      onOpenChange={(open) => !open && setQuery('')}
      trigger={
        <button
          type="button"
          className={cn('filter-chip', active && 'filter-chip-active')}
          data-testid={`filter-${label.toLowerCase().replace(/\s+/g, '-')}`}
        >
          {icon}
          {text}
          <ChevronDown />
        </button>
      }
    >
      {(close) => (
        <div style={{ width: 260 }}>
          {options.length > 7 && (
            <div className="picker-search">
              <input
                type="search"
                className="field-control"
                style={{ height: 32 }}
                placeholder={`Filter ${label.toLowerCase()}`}
                value={query}
                onChange={(e) => setQuery(e.target.value)}
                aria-label={`Filter ${label} options`}
                autoFocus
              />
            </div>
          )}
          <div className="picker-list" role="listbox" aria-multiselectable={!single} aria-label={label}>
            {filtered.length === 0 && <div className="picker-empty">No matches</div>}
            {filtered.map((o) => {
              const selected = value.includes(o.value);
              return (
                <button
                  key={o.value}
                  type="button"
                  role="option"
                  aria-selected={selected}
                  className="picker-option"
                  onClick={() => toggle(o.value, close)}
                >
                  <span className="picker-check" aria-hidden="true">
                    {selected && <Check size={16} />}
                  </span>
                  {o.color && (
                    <span className="chip-dot" style={{ background: o.color }} aria-hidden="true" />
                  )}
                  <span className="truncate">{o.label}</span>
                </button>
              );
            })}
          </div>
          {active && (
            <div style={{ padding: 8, borderTop: '1px solid var(--color-border)' }}>
              <button
                type="button"
                className="link-button row"
                onClick={() => {
                  onChange([]);
                  close();
                }}
              >
                <X size={14} /> Clear
              </button>
            </div>
          )}
        </div>
      )}
    </Popover>
  );
}
