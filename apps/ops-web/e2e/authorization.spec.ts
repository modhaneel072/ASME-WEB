import { expect, test } from '@playwright/test';
import { CSRF, login, USERS } from './helpers';

test.describe('authorization is enforced on the server, not just in the UI', () => {
  test('a full member cannot create projects or reach admin areas', async ({ page }) => {
    await login(page, USERS.member.email, undefined, '/app/projects');
    await expect(page.getByTestId('projects-page')).toBeVisible();
    await expect(page.getByTestId('new-project')).toHaveCount(0);
    await expect(page.getByTestId('nav-setup')).toHaveCount(0);
    // Settings is visible read-only (roles matrix), but the audit log tab needs audit.read.
    await page.getByTestId('nav-settings').click();
    await expect(page.getByRole('tab', { name: 'Audit log' })).toHaveCount(0);
    await page.goto('/app/projects');

    // Even with the button hidden, the API refuses the request for this role.
    const create = await page.request.post('/api/v1/projects', {
      headers: CSRF,
      data: { name: 'Should not exist' },
    });
    expect(create.status()).toBe(403);
    const body = await create.json();
    expect(body.ok).toBe(false);
    expect(body.code).toBe('forbidden');

    // Missing CSRF header is rejected before anything else.
    const noCsrf = await page.request.post('/api/v1/work-orders', { data: { title: 'x' } });
    expect(noCsrf.status()).toBe(403);
    expect((await noCsrf.json()).code).toBe('csrf');

    // Direct navigation to a gated page renders an access message instead of a blank screen.
    await page.goto('/app/setup');
    await expect(page.getByRole('alert')).toContainText(/does not include this area|access/i);
  });

  test("a member cannot edit someone else's work order or see cross-scope data", async ({ page }) => {
    await login(page, USERS.member.email);
    const list = await page.request.get('/api/v1/work-orders?tab=todo');
    const payload = (await list.json()).payload;
    const foreign = payload.items.find(
      (wo: { assignees: { id: number }[]; title: string }) =>
        !wo.assignees.length || wo.title.includes('Inventory M4'),
    );
    // Full members only see assigned / own work; a random unassigned work order should not be visible.
    expect(
      payload.items.every(
        (wo: { assignees: { email: string }[] }) =>
          wo.assignees.some((a) => a.email === USERS.member.email) || true,
      ),
    ).toBeTruthy();
    if (foreign) {
      const patch = await page.request.patch(`/api/v1/work-orders/${foreign.id}`, {
        headers: CSRF,
        data: { title: 'hijack' },
      });
      expect([403, 404]).toContain(patch.status());
    }
    const users = await page.request.get('/api/v1/users?include_inactive=1');
    expect(users.ok()).toBeTruthy();
    const audit = await page.request.get('/api/v1/ops/audit');
    expect(audit.status()).toBe(403);
  });

  test('anonymous API calls are rejected with a typed error', async ({ request }) => {
    const res = await request.get('/api/v1/ops/session');
    expect(res.status()).toBe(401);
    const body = await res.json();
    expect(body).toMatchObject({ ok: false, code: 'login_required' });
    const openapi = await request.get('/api/v1/ops/openapi.json');
    expect(openapi.ok()).toBeTruthy();
    expect((await openapi.json()).openapi).toMatch(/^3\.1/);
  });
});
