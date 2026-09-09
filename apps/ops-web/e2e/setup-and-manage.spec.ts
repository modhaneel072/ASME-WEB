import { expect, test } from '@playwright/test';
import { login, unique, USERS } from './helpers';

test.describe('setup center and manage pages', () => {
  test('setup center reflects persisted configuration and never advertises unbuilt modules', async ({
    page,
  }) => {
    await login(page, USERS.admin.email, undefined, '/app/setup');
    await expect(page.getByTestId('setup-center')).toBeVisible();
    // Demo data already satisfies locations, assets, teams/users, project and categories.
    for (const key of ['locations', 'assets', 'teams_users', 'first_project', 'categories'])
      await expect(page.getByTestId(`setup-task-${key}`)).toContainText('Done');
    // Modules that do not exist yet are labelled, with no call to action.
    for (const key of [
      'parts',
      'procedure',
      'maintenance_plan',
      'request_portal',
      'automation',
      'dashboard',
    ]) {
      const row = page.getByTestId(`setup-task-${key}`);
      await expect(row).toContainText('Not in this release');
      await expect(row.getByRole('link')).toHaveCount(0);
      await expect(row.getByRole('button')).toHaveCount(0);
    }
    // Sidebar lists only built modules.
    for (const id of [
      'nav-setup',
      'nav-work-orders',
      'nav-projects',
      'nav-reporting',
      'nav-assets',
      'nav-locations',
      'nav-categories',
      'nav-teams-users',
      'nav-settings',
    ])
      await expect(page.getByTestId(id)).toBeVisible();
    await expect(page.getByTestId('sidebar')).not.toContainText(/Requests|Parts|Procedures|Automations/);

    // Completing the chapter profile advances the checklist.
    const before = await page.getByTestId('setup-task-chapter_profile').textContent();
    if (!before?.includes('Done')) {
      await page.getByTestId('setup-task-chapter_profile').getByRole('link').click();
      await expect(page.getByTestId('chapter-form')).toBeVisible();
      await page.getByTestId('chapter-save').click();
      await page.goto('/app/setup');
      await expect(page.getByTestId('setup-task-chapter_profile')).toContainText('Done');
    }
    // Reading the guide completes the optional task.
    await page.getByRole('tab', { name: 'Officer guide' }).click();
    const guide = page.getByTestId('guide-read');
    if (await guide.isEnabled()) await guide.click();
    await expect(guide).toBeDisabled();
  });

  test('admin manages teams, users, locations and categories', async ({ page }) => {
    await login(page, USERS.admin.email, undefined, '/app/teams-users');
    await page.getByTestId('new-team').click();
    const teamName = unique('Test Stand Crew');
    await page.getByTestId('team-name').fill(teamName);
    await page.locator('#team-members').click();
    await page.getByRole('option', { name: new RegExp(USERS.member2.name) }).click();
    await page.keyboard.press('Escape');
    await page.getByRole('checkbox', { name: new RegExp(USERS.member2.name) }).check();
    await page.getByTestId('team-submit').click();
    await expect(page.getByTestId('team-row').filter({ hasText: teamName })).toContainText(
      USERS.member2.name,
    );

    await page.getByRole('tab', { name: 'Users' }).click();
    await page.getByTestId('invite-user').click();
    const email = `e2e-${Date.now().toString(36)}@uiowa.edu`;
    await page.getByTestId('invite-name').fill('E2E Invitee');
    await page.getByTestId('invite-email').fill(email);
    await page.getByTestId('invite-submit').click();
    await expect(page.getByTestId('user-row').filter({ hasText: email })).toHaveCount(1);

    await page.getByTestId('nav-locations').click();
    await page.getByTestId('new-location').click();
    const locName = unique('Paint Booth');
    await page.getByTestId('location-name').fill(locName);
    await page.getByLabel('Room').fill('1250');
    await page.getByTestId('location-submit').click();
    await expect(page.getByTestId('location-row').filter({ hasText: locName })).toContainText('1250');

    await page.getByTestId('nav-categories').click();
    await page.getByTestId('new-category').click();
    const catName = unique('Composites');
    await page.getByTestId('category-name').fill(catName);
    await page.getByTestId('category-submit').click();
    await expect(page.getByTestId('category-row').filter({ hasText: catName })).toHaveCount(1);
    // Duplicate names are rejected by the server and surfaced in the form.
    await page.getByTestId('new-category').click();
    await page.getByTestId('category-name').fill(catName);
    await page.getByTestId('category-submit').click();
    await expect(page.getByTestId('category-form').getByRole('alert')).toContainText(/already exists/i);
  });

  test('assets: detail, status change and the operations dashboard render real numbers', async ({ page }) => {
    await login(page, USERS.admin.email, undefined, '/app/assets');
    await page.getByTestId('asset-row').filter({ hasText: '3D Printer 01' }).click();
    await expect(page.getByTestId('asset-detail')).toContainText('PRN-01');
    await page.getByTestId('asset-status').click();
    await page.getByLabel('Status', { exact: true }).selectOption('OFFLINE_PLANNED');
    await page.getByLabel('Reason').fill('Nozzle swap');
    await page.getByTestId('asset-status-submit').click();
    await expect(page.getByTestId('asset-detail')).toContainText('Offline (planned)');
    await expect(page.getByTestId('asset-detail')).toContainText('Nozzle swap');

    await page.getByTestId('nav-reporting').click();
    await expect(page.getByTestId('kpi-grid')).toBeVisible();
    const open = await page.getByTestId('kpi-open').textContent();
    expect(Number(open)).toBeGreaterThan(0);
    await page.getByRole('button', { name: '7 days' }).click();
    await expect(page.getByTestId('kpi-grid')).toBeVisible();
  });
});
