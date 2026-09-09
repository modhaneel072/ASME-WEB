import { expect, test } from '@playwright/test';
import { login, unique, USERS } from './helpers';

test.describe('projects', () => {
  test('admin creates a project, assigns a lead and adds a milestone', async ({ page }) => {
    await login(page, USERS.admin.email, undefined, '/app/projects');
    await page.getByTestId('new-project').click();
    const form = page.getByTestId('project-form');
    await expect(form).toBeVisible();
    const name = unique('Wind Tunnel Rebuild');
    await page.getByTestId('project-name').fill(name);
    await form
      .getByLabel('Code', { exact: true })
      .fill(`WT${Date.now().toString(36).slice(-4).toUpperCase()}`);
    await form.locator('#project-lead').click();
    await page.getByRole('option', { name: new RegExp(USERS.projectLead.name) }).click();
    await form.getByLabel('Target date').fill('2027-04-30');
    await page.getByTestId('project-submit').click();

    await expect(page.getByTestId('project-detail')).toBeVisible();
    await expect(page.getByRole('heading', { level: 1 })).toHaveText(name);
    await expect(page.getByTestId('project-detail')).toContainText(`Lead: ${USERS.projectLead.name}`);

    await page.getByRole('tab', { name: /Milestones/ }).click();
    await page.getByTestId('add-milestone').click();
    await page.getByTestId('milestone-name').fill('Tunnel mounted');
    await page.getByLabel('Due date').fill('2027-02-01');
    await page.getByTestId('milestone-submit').click();
    await expect(page.getByTestId('milestone-row')).toHaveCount(1);
    await expect(page.getByTestId('milestone-row').first()).toContainText('Tunnel mounted');

    await page.getByRole('tab', { name: /Members/ }).click();
    await expect(page.getByTestId('member-row').first()).toContainText(USERS.projectLead.name);
  });

  test('project work-orders tab lists only that project and opens the create pane with the project preset', async ({
    page,
  }) => {
    await login(page, USERS.projectLead.email, undefined, '/app/projects');
    await page.getByTestId('project-card').filter({ hasText: 'Crater Cruncher Rover' }).click();
    await page.getByRole('tab', { name: /Work orders/ }).click();
    await expect(page.getByTestId('work-order-row').first()).toBeVisible();
    const rows = page.getByTestId('work-order-row');
    const count = await rows.count();
    for (let i = 0; i < Math.min(count, 5); i++)
      await expect(rows.nth(i)).toContainText('Crater Cruncher Rover');
    await page.getByTestId('new-work-order').click();
    const form = page.getByTestId('work-order-form');
    await expect(form).toBeVisible();
    await expect(form.locator('#wo-project')).toContainText('Crater Cruncher Rover');
  });
});
