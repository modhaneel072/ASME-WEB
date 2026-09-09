import { render, screen } from '@testing-library/react';
import userEvent from '@testing-library/user-event';
import { describe, expect, it, vi } from 'vitest';
import { FilterChip } from './FilterChip';

const options = [
  { value: 'OPEN', label: 'Open' },
  { value: 'IN_PROGRESS', label: 'In progress' },
  { value: 'ON_HOLD', label: 'On hold' },
];

describe('FilterChip', () => {
  it('shows the label when inactive and the selection when active', async () => {
    const onChange = vi.fn();
    const { rerender } = render(
      <FilterChip label="Status" options={options} value={[]} onChange={onChange} />,
    );
    const chip = screen.getByRole('button', { name: /status/i });
    expect(chip).not.toHaveClass('filter-chip-active');
    await userEvent.click(chip);
    await userEvent.click(screen.getByRole('option', { name: 'Open' }));
    expect(onChange).toHaveBeenCalledWith(['OPEN']);
    rerender(<FilterChip label="Status" options={options} value={['OPEN']} onChange={onChange} />);
    expect(screen.getByRole('button', { name: /status: open/i })).toHaveClass('filter-chip-active');
  });

  it('supports multi-select toggling and clearing', async () => {
    const onChange = vi.fn();
    render(<FilterChip label="Status" options={options} value={['OPEN']} onChange={onChange} />);
    await userEvent.click(screen.getByRole('button', { name: /status/i }));
    await userEvent.click(screen.getByRole('option', { name: 'On hold' }));
    expect(onChange).toHaveBeenLastCalledWith(['OPEN', 'ON_HOLD']);
    await userEvent.click(screen.getByRole('option', { name: 'Open' }));
    expect(onChange).toHaveBeenLastCalledWith([]);
    await userEvent.click(screen.getByRole('button', { name: /clear/i }));
    expect(onChange).toHaveBeenLastCalledWith([]);
  });

  it('closes on Escape', async () => {
    render(<FilterChip label="Status" options={options} value={[]} onChange={() => {}} />);
    await userEvent.click(screen.getByRole('button', { name: /status/i }));
    expect(screen.getByRole('listbox')).toBeInTheDocument();
    await userEvent.keyboard('{Escape}');
    expect(screen.queryByRole('listbox')).not.toBeInTheDocument();
  });
});
