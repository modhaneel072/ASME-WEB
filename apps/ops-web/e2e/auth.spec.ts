import { expect, test } from '@playwright/test';
import { login, logout, PASSWORD, USERS } from './helpers';

test.describe('authentication', () => {
  test('redirects anonymous visitors to the login page and back after signing in', async ({ page }) => {
    await page.goto('/app/projects');
    await expect(page).toHaveURL(/\/app\/login\?next=%2Fprojects/);
    await page.getByTestId('login-identifier').fill(USERS.admin.email);
    await page.getByTestId('login-password').fill(PASSWORD);
    await page.getByTestId('login-submit').click();
    await expect(page).toHaveURL(/\/app\/projects$/);
    await expect(page.getByTestId('projects-page')).toBeVisible();
  });

  test('signs in with a username as well as an email', async ({ page }) => {
    await login(page, USERS.teamLead.username);
    await expect(page.getByTestId('user-menu')).toContainText(USERS.teamLead.name);
    await logout(page);
    await login(page, USERS.teamLead.email);
    await expect(page.getByTestId('user-menu')).toContainText(USERS.teamLead.name);
  });

  test('shows a clear error for a wrong password without leaking which part was wrong', async ({ page }) => {
    await page.goto('/app/login');
    await page.getByTestId('login-identifier').fill(USERS.admin.email);
    await page.getByTestId('login-password').fill('definitely-wrong');
    await page.getByTestId('login-submit').click();
    await expect(page.getByTestId('login-error')).toContainText('Invalid email/username or password');
    await expect(page).toHaveURL(/\/app\/login/);
  });

  test('legacy and public routes keep working alongside the new app', async ({ page, request }) => {
    const health = await request.get('/healthz');
    expect(health.ok()).toBeTruthy();
    const publicHome = await request.get('/');
    expect(publicHome.status()).toBe(200);
    const legacy = await request.get('/legacy/app', { maxRedirects: 0 });
    expect([200, 301, 302, 303, 307, 308]).toContain(legacy.status());
    const alias = await request.get('/auth/login', { maxRedirects: 0 });
    expect(alias.status()).toBeGreaterThanOrEqual(300);
    expect(alias.headers()['location']).toContain('/app/login');
    await page.goto('/app/does-not-exist');
    await expect(page).toHaveURL(/\/app\/login/);
  });
});
