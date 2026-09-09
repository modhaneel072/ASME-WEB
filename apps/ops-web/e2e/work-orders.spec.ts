import path from 'node:path';
import { fileURLToPath } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
import { expect, test } from '@playwright/test';
import { login, logout, unique, USERS } from './helpers';

test.describe.serial('work-order vertical slice', () => {
  const title = unique('Torque check on RL wheel module');

  test('team lead creates a work order in the right-side pane and assigns a member', async ({ page }) => {
    await login(page, USERS.teamLead.email);
    await expect(page.getByTestId('work-orders-page')).toBeVisible();
    await page.getByTestId('new-work-order').click();
    const form = page.getByTestId('work-order-form');
    await expect(form).toBeVisible();
    // The pane sits on the right edge and leaves the list visible behind it.
    const box = await form.boundingBox();
    const viewport = page.viewportSize()!;
    expect(box!.x + box!.width).toBeGreaterThan(viewport.width - 2);
    expect(box!.width).toBeLessThan(viewport.width * 0.6);

    await page.getByTestId('wo-title').fill(title);
    await form.getByLabel('Description').fill('Verify hub fastener torque after the v4 hub swap.');
    await form.getByRole('button', { name: 'High' }).click();
    await form.locator('#wo-project').click();
    await page.getByRole('option', { name: /Crater Cruncher Rover/ }).click();
    await form.locator('#wo-assignees').click();
    await page.getByRole('option', { name: new RegExp(USERS.member.name) }).click();
    await page.keyboard.press('Escape');
    const due = new Date(Date.now() + 3 * 86400000);
    await page.getByTestId('wo-due').fill(`${due.toISOString().slice(0, 10)}T17:00`);
    await page.getByTestId('work-order-submit').click();

    await expect(page.getByTestId('work-order-form')).toBeHidden();
    const detail = page.getByTestId('work-order-detail');
    await expect(detail).toBeVisible();
    await expect(page.getByTestId('detail-title')).toHaveText(title);
    await expect(detail).toContainText('Open');
    await expect(detail).toContainText(USERS.member.name);
    await expect(page.getByTestId('work-order-row').filter({ hasText: title })).toHaveCount(1);
  });

  test('member filters to their work, starts it, comments, uploads, logs time and completes it', async ({
    page,
  }) => {
    await login(page, USERS.member.email);
    await page.getByTestId('filter-assignee').click();
    await page.getByRole('option', { name: 'Assigned to me' }).click();
    await page.keyboard.press('Escape');
    await expect(page).toHaveURL(/filter%5Bassignee%5D=me|filter\[assignee\]=me/);
    const row = page.getByTestId('work-order-row').filter({ hasText: title });
    await expect(row).toHaveCount(1);
    await row.click();

    const detail = page.getByTestId('work-order-detail');
    await expect(page.getByTestId('detail-title')).toHaveText(title);
    await page.getByTestId('action-start').click();
    await expect(detail.getByText('In progress').first()).toBeVisible();

    await page.getByTestId('comment-input').fill('Torqued to 12 Nm, all four bolts.');
    await page.getByTestId('comment-submit').click();
    await expect(page.getByTestId('comment').filter({ hasText: 'Torqued to 12 Nm' })).toHaveCount(1);

    await page.getByTestId('file-input').setInputFiles(path.join(here, 'fixtures', 'torque-sheet.txt'));
    await expect(page.getByTestId('file-row').filter({ hasText: 'torque-sheet.txt' })).toHaveCount(1);

    await page.getByTestId('log-time').click();
    await page.getByTestId('time-minutes').fill('25');
    await page.getByTestId('time-submit').click();
    await expect(page.getByTestId('time-cost-section')).toContainText('25m');

    await page.getByTestId('action-complete').click();
    await page.getByTestId('complete-note').fill('All fasteners within spec.');
    await page.getByTestId('complete-submit').click();
    await expect(detail.getByText('Done').first()).toBeVisible();
    await expect(detail).toContainText('All fasteners within spec.');

    // It moved from To Do to Done.
    await page.getByRole('tab', { name: /Done/ }).click();
    await expect(page.getByTestId('work-order-row').filter({ hasText: title })).toHaveCount(1);
    await logout(page);
  });

  test('sub-work orders, saved views and the table layout', async ({ page }) => {
    await login(page, USERS.teamLead.email);
    await page.getByTestId('work-order-row').filter({ hasText: 'Wheel hub CAD revision (v4)' }).click();
    await expect(page.getByTestId('sub-work-orders')).toContainText('done');
    await expect(page.getByTestId('sub-row')).toHaveCount(3);
    await page.getByTestId('add-sub').click();
    await page.getByTestId('sub-title').fill('Order v4 hub fasteners');
    await page.getByTestId('sub-submit').click();
    await expect(page.getByTestId('sub-row')).toHaveCount(4);

    await page.getByTestId('filter-priority').click();
    await page.getByRole('option', { name: 'High' }).click();
    await page.keyboard.press('Escape');
    await page.getByTestId('saved-views').click();
    await page.getByTestId('save-view').click();
    const viewName = unique('High priority');
    await page.getByLabel('Name').fill(viewName);
    await page.getByTestId('save-view-submit').click();
    await page.getByTestId('clear-filters').click();
    await page.getByTestId('saved-views').click();
    await page.getByTestId('saved-view-item').filter({ hasText: viewName }).click();
    await expect(page.getByTestId('filter-priority')).toContainText('Priority: High');

    await page.getByRole('button', { name: 'Table view' }).click();
    await expect(page.getByTestId('work-order-table')).toBeVisible();
  });
});
