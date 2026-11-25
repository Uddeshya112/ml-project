const { test, expect } = require('@playwright/test');

test('theme, custom JD, and reset controls work', async ({ page }) => {
  await page.goto('/', { waitUntil: 'domcontentloaded' });

  await page.getByRole('checkbox', { name: 'Use preset role' }).uncheck();
  await page.locator('#jdTextarea').fill('Python developer with Docker and Kubernetes');
  await page.getByRole('checkbox', { name: 'Include sample corpus (5)' }).uncheck();
  await page.getByRole('checkbox', { name: 'Include sample corpus (5)' }).check();

  await page.getByRole('button', { name: 'Run TF-IDF & Cosine Ranking' }).click();
  await expect(page.getByText('Candidate Match Intelligence')).toBeVisible();

  await page.getByRole('button', { name: 'Toggle theme' }).click();
  await expect(page.locator('html')).toHaveAttribute('data-theme', 'light');

  await page.getByRole('button', { name: 'Reset Workspace' }).click();
  await expect(page.getByRole('checkbox', { name: 'Use preset role' })).toBeChecked();
  await expect(page.getByRole('checkbox', { name: 'Include sample corpus (5)' })).toBeChecked();
  await expect(page.locator('#thresholdValueLabel')).toHaveText('0%');
});
