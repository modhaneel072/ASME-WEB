import type { ReactNode } from 'react';
import {
  AlertTriangle,
  Ban,
  CheckCircle2,
  CircleDashed,
  CirclePause,
  CirclePlay,
  Flag,
  Tag,
} from 'lucide-react';
import { cn } from '@/lib/cn';
import { STATUS_LABEL, humanize } from '@/lib/format';
import type { CategoryChip, WoPriority, WoStatus } from '@/contracts/types';

export function Badge({
  tone = 'neutral',
  children,
  className,
  icon,
}: {
  tone?: string;
  children: ReactNode;
  className?: string;
  icon?: ReactNode;
}) {
  return (
    <span className={cn('badge', `badge-${tone}`, className)}>
      {icon}
      {children}
    </span>
  );
}

const STATUS_ICON: Record<WoStatus, ReactNode> = {
  DRAFT: <CircleDashed />,
  OPEN: <CircleDashed />,
  IN_PROGRESS: <CirclePlay />,
  ON_HOLD: <CirclePause />,
  DONE: <CheckCircle2 />,
  CANCELED: <Ban />,
  SKIPPED: <Ban />,
};

export function StatusBadge({ status }: { status: WoStatus | string }) {
  const key = status.toLowerCase();
  return (
    <Badge tone={key} icon={STATUS_ICON[status as WoStatus]} className="status-badge">
      {STATUS_LABEL[status] || humanize(status)}
    </Badge>
  );
}

export function PriorityBadge({ priority }: { priority: WoPriority | string }) {
  if (priority === 'NONE') return null;
  return (
    <Badge tone={priority.toLowerCase()} icon={priority === 'CRITICAL' ? <AlertTriangle /> : <Flag />}>
      {humanize(priority)}
    </Badge>
  );
}

export function CategoryChipView({ category }: { category: CategoryChip | { name: string; color: string } }) {
  return (
    <span className="chip" title={category.name}>
      <span className="chip-dot" style={{ background: category.color }} aria-hidden="true" />
      {category.name}
    </span>
  );
}

export function OverdueBadge() {
  return (
    <Badge tone="overdue" icon={<AlertTriangle />}>
      Overdue
    </Badge>
  );
}

export function BlockedBadge() {
  return <Badge tone="blocked">Blocked</Badge>;
}

export function GenericChip({ label, color }: { label: string; color?: string | null }) {
  return (
    <span className="chip">
      {color ? (
        <span className="chip-dot" style={{ background: color }} aria-hidden="true" />
      ) : (
        <Tag size={12} />
      )}
      {label}
    </span>
  );
}
