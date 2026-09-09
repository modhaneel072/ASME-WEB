import { useEffect, useState, type ReactNode } from 'react';
import { Link } from 'react-router-dom';
import { ChevronRight, Search, X } from 'lucide-react';
import { IconButton } from './Button';

export function PageHeader({
  title,
  subtitle,
  actions,
  breadcrumbs,
  testId,
}: {
  title: ReactNode;
  subtitle?: ReactNode;
  actions?: ReactNode;
  breadcrumbs?: { label: string; to?: string }[];
  testId?: string;
}) {
  return (
    <div className="page-head" data-testid={testId}>
      <div className="page-title-wrap">
        {breadcrumbs && breadcrumbs.length > 0 && (
          <nav className="breadcrumbs" aria-label="Breadcrumb">
            {breadcrumbs.map((b, i) => (
              <span key={i} className="row" style={{ gap: 6 }}>
                {b.to ? <Link to={b.to}>{b.label}</Link> : <span>{b.label}</span>}
                {i < breadcrumbs.length - 1 && <ChevronRight size={12} />}
              </span>
            ))}
          </nav>
        )}
        <h1 className="page-title">{title}</h1>
        {subtitle && <div className="page-subtitle">{subtitle}</div>}
      </div>
      {actions && <div className="page-actions">{actions}</div>}
    </div>
  );
}

/** Debounced search input that writes through to the URL-backed list state. */
export function SearchField({
  value,
  onChange,
  placeholder = 'Search',
  label = 'Search',
  delay = 250,
}: {
  value: string;
  onChange: (v: string) => void;
  placeholder?: string;
  label?: string;
  delay?: number;
}) {
  const [draft, setDraft] = useState(value);
  useEffect(() => setDraft(value), [value]);
  useEffect(() => {
    if (draft === value) return;
    const t = window.setTimeout(() => onChange(draft), delay);
    return () => window.clearTimeout(t);
  }, [draft, value, onChange, delay]);
  return (
    <div className="search">
      <Search aria-hidden="true" />
      <input
        type="search"
        className="field-control"
        value={draft}
        onChange={(e) => setDraft(e.target.value)}
        placeholder={placeholder}
        aria-label={label}
      />
      {draft && (
        <IconButton label="Clear search" size="sm" className="search-clear" onClick={() => setDraft('')}>
          <X />
        </IconButton>
      )}
    </div>
  );
}
