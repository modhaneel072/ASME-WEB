import type { ReactNode } from 'react';
import { Link } from 'react-router-dom';

export interface TabDef<T extends string = string> {
  key: T;
  label: ReactNode;
  count?: number | null;
  to?: string;
}

export function Tabs<T extends string>({
  tabs,
  value,
  onChange,
  label,
}: {
  tabs: TabDef<T>[];
  value: T;
  onChange?: (key: T) => void;
  label: string;
}) {
  return (
    <div className="tabs" role="tablist" aria-label={label}>
      {tabs.map((tab) =>
        tab.to ? (
          <Link key={tab.key} to={tab.to} className="tab" role="tab" aria-selected={tab.key === value}>
            {tab.label}
            {tab.count !== undefined && tab.count !== null && <span className="tab-count">{tab.count}</span>}
          </Link>
        ) : (
          <button
            key={tab.key}
            type="button"
            className="tab"
            role="tab"
            aria-selected={tab.key === value}
            onClick={() => onChange?.(tab.key)}
          >
            {tab.label}
            {tab.count !== undefined && tab.count !== null && <span className="tab-count">{tab.count}</span>}
          </button>
        ),
      )}
    </div>
  );
}
