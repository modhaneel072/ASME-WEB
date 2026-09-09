import { Compass } from 'lucide-react';
import { EmptyState, LinkButton } from '@/ui';

export function NotFoundPage() {
  return (
    <div className="page">
      <div className="page-body" style={{ paddingTop: 48 }}>
        <EmptyState
          icon={<Compass />}
          title="That page does not exist"
          body="It may have moved, or the link was mistyped. Modules that are not built yet are not linked anywhere on purpose."
          actions={
            <LinkButton to="/work-orders" variant="primary">
              Go to work orders
            </LinkButton>
          }
        />
      </div>
    </div>
  );
}
