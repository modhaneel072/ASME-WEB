import { render, screen } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import { PriorityBadge, StatusBadge } from './Badge';

describe('badges', () => {
  it('renders readable status labels with a tone class', () => {
    render(<StatusBadge status="IN_PROGRESS" />);
    const badge = screen.getByText('In progress');
    expect(badge.closest('.badge')).toHaveClass('badge-in_progress');
  });

  it('hides the NONE priority and shows others', () => {
    const { container } = render(<PriorityBadge priority="NONE" />);
    expect(container).toBeEmptyDOMElement();
    render(<PriorityBadge priority="CRITICAL" />);
    expect(screen.getByText('Critical').closest('.badge')).toHaveClass('badge-critical');
  });
});
