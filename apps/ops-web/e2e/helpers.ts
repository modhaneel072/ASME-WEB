import { expect, type Page } from '@playwright/test';

/** Development-only seed credentials (asme/ops/seeds.py). Never used in production. */
export const PASSWORD = 'ChangeMe123!';
export const USERS = {
  admin: { email: 'admin@uiowa.edu', username: 'admin', name: 'Ada Admin' },
  projectLead: { email: 'pnatarajan@uiowa.edu', username: 'pnatarajan', name: 'Priya Natarajan' },
  teamLead: { email: 'mbell@uiowa.edu', username: 'mbell', name: 'Marcus Bell' },
  member: { email: 'avery@uiowa.edu', username: 'avery', name: 'Avery Johnson' },
  member2: { email: 'taylor@uiowa.edu', username: 'taylor', name: 'Taylor Kim' },
};

export const CSRF = { 'X-Requested-With': 'ASME-Ops' };

export async function login(page: Page, identifier: string, password = PASSWORD, next = '/app/work-orders') {
  await page.goto(`/app/login?next=${encodeURIComponent(next.replace(/^\/app/, ''))}`);
  await page.getByTestId('login-identifier').fill(identifier);
  await page.getByTestId('login-password').fill(password);
  await page.getByTestId('login-submit').click();
  await expect(page.getByTestId('sidebar')).toBeVisible();
}

export async function logout(page: Page) {
  await page.getByTestId('user-menu').click();
  await page.getByRole('menuitem', { name: 'Sign out' }).click();
  await expect(page.getByTestId('login-form')).toBeVisible();
}

export function unique(prefix: string) {
  return `${prefix} ${Date.now().toString(36)}`;
}
