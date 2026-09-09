import AxeBuilder from '@axe-core/playwright';
import { expect, test } from '@playwright/test';
import { login, USERS } from './helpers';

// Runs in every Playwright project (1440×900, 1366×768, 768×1024, 390×844) via @responsive.
test.describe('@responsive layout and accessibility', () => {
  test('work orders list, detail and create pane adapt to the viewport', async ({ page }, testInfo) => {
    await login(page, USERS.teamLead.email);
    const width = page.viewportSize()!.width;
    if (width < 900) {
      await expect(page.getByTestId('open-nav')).toBeVisible();
      await page.getByTestId('open-nav').click();
      await expect(page.getByTestId('sidebar')).toBeVisible();
      await page.keyboard.press('Escape');
      await page.getByTestId('open-nav').click();
      await page.getByTestId('nav-work-orders').click();
    } else {
      await expect(page.getByTestId('open-nav')).toBeHidden();
    }
    await expect(page.getByTestId('work-order-row').first()).toBeVisible();
    await testInfo.attach(`work-orders-list-${width}`, {
      body: await page.screenshot({ fullPage: false, animations: 'disabled' }),
      contentType: 'image/png',
    });

    await page.getByTestId('work-order-row').first().click();
    await expect(page.getByTestId('work-order-detail')).toBeVisible();
    if (width < 1050) {
      // Compact: the detail replaces the list and offers a back button.
      await expect(page.getByTestId('detail-back')).toBeVisible();
      await expect(page.getByTestId('work-order-list')).toBeHidden();
    } else {
      await expect(page.getByTestId('detail-close')).toBeVisible();
      await expect(page.getByTestId('work-order-list')).toBeVisible();
    }
    await testInfo.attach(`work-order-detail-${width}`, {
      body: await page.screenshot({ animations: 'disabled' }),
      contentType: 'image/png',
    });
    // Body never scrolls horizontally.
    const overflow = await page.evaluate(
      () => document.documentElement.scrollWidth - document.documentElement.clientWidth,
    );
    expect(overflow).toBeLessThanOrEqual(1);

    await page.getByTestId('new-work-order').click();
    await expect(page.getByTestId('work-order-form')).toBeVisible();
    await page.waitForTimeout(250); // let the 180ms slide-in finish so the capture is opaque
    await testInfo.attach(`work-order-create-${width}`, {
      body: await page.screenshot({ animations: 'disabled' }),
      contentType: 'image/png',
    });
    await page.keyboard.press('Escape');
    await expect(page.getByTestId('work-order-form')).toBeHidden();
  });

  test('key screens pass automated accessibility checks', async ({ page }, testInfo) => {
    const violationsFor = async (name: string) => {
      const results = await new AxeBuilder({ page })
        .withTags(['wcag2a', 'wcag2aa'])
        .disableRules(['color-contrast'])
        .analyze();
      await testInfo.attach(`axe-${name}-${page.viewportSize()!.width}`, {
        body: JSON.stringify(results.violations, null, 2),
        contentType: 'application/json',
      });
      return results.violations.filter((v) => v.impact === 'critical' || v.impact === 'serious');
    };
    await page.goto('/app/login');
    expect(await violationsFor('login')).toEqual([]);
    await login(page, USERS.admin.email, undefined, '/app/setup');
    await expect(page.getByTestId('setup-center')).toBeVisible();
    expect(await violationsFor('setup-center')).toEqual([]);
    await page.goto('/app/work-orders');
    await expect(page.getByTestId('work-order-row').first()).toBeVisible();
    expect(await violationsFor('work-orders')).toEqual([]);
    await page.goto('/app/reporting');
    await expect(page.getByTestId('kpi-grid')).toBeVisible();
    expect(await violationsFor('reporting')).toEqual([]);
  });
});
