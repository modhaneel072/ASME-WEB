import { cn } from '@/lib/cn';
import type { Ref, UserRef } from '@/contracts/types';

const PALETTE = [
  '#eaf4ff|#05589a',
  '#fff8d7|#7a5b00',
  '#e7f8f2|#007a57',
  '#efeafd|#5b3fc4',
  '#fff4dd|#a45e00',
  '#fff0f0|#b53a3a',
];

function colorsFor(seed: number | string): { background: string; color: string } {
  const n = typeof seed === 'number' ? seed : Array.from(seed).reduce((a, c) => a + c.charCodeAt(0), 0);
  const [background, color] = PALETTE[Math.abs(n) % PALETTE.length].split('|');
  return { background, color };
}

export function Avatar({
  user,
  size = 'md',
  className,
}: {
  user: Pick<UserRef, 'id' | 'name' | 'initials'>;
  size?: 'sm' | 'md' | 'lg';
  className?: string;
}) {
  return (
    <span
      className={cn('avatar', size !== 'md' && `avatar-${size}`, className)}
      style={colorsFor(user.id)}
      title={user.name}
      aria-label={user.name}
      role="img"
    >
      {user.initials}
    </span>
  );
}

export function TeamAvatar({ team, size = 'md' }: { team: Ref; size?: 'sm' | 'md' | 'lg' }) {
  const initials = team.name
    .split(/\s+/)
    .slice(0, 2)
    .map((p) => p[0]?.toUpperCase() || '')
    .join('');
  return (
    <span
      className={cn('avatar avatar-team', size !== 'md' && `avatar-${size}`)}
      title={`${team.name} (team)`}
      aria-label={`${team.name} team`}
      role="img"
    >
      {initials}
    </span>
  );
}

export function AvatarStack({
  users,
  teams = [],
  max = 3,
  size = 'sm',
}: {
  users: UserRef[];
  teams?: Ref[];
  max?: number;
  size?: 'sm' | 'md';
}) {
  const total = users.length + teams.length;
  if (total === 0) return <span className="text-muted text-caption">Unassigned</span>;
  const shownUsers = users.slice(0, max);
  const remainingSlots = Math.max(0, max - shownUsers.length);
  const shownTeams = teams.slice(0, remainingSlots);
  const overflow = total - shownUsers.length - shownTeams.length;
  return (
    <span
      className="avatar-stack"
      aria-label={[...users.map((u) => u.name), ...teams.map((t) => `${t.name} team`)].join(', ')}
    >
      {shownUsers.map((u) => (
        <Avatar key={`u${u.id}`} user={u} size={size} />
      ))}
      {shownTeams.map((t) => (
        <TeamAvatar key={`t${t.id}`} team={t} size={size} />
      ))}
      {overflow > 0 && (
        <span className={cn('avatar avatar-more', size !== 'md' && `avatar-${size}`)}>+{overflow}</span>
      )}
    </span>
  );
}
