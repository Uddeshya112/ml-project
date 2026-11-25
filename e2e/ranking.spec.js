const { test, expect } = require('@playwright/test');

test('sample corpus produces and filters rankings', async ({ page }) => {
  await page.goto('/', { waitUntil: 'domcontentloaded' });
  await expect(page.getByRole('heading', { name: 'Rank every candidate, with mathematical clarity.' })).toBeVisible();

  await page.getByRole('button', { name: 'Run TF-IDF & Cosine Ranking' }).click();
  await expect(page.getByText('Documents Ranked')).toBeVisible();
  await expect(page.getByText('devops_engineer_resume.txt').first()).toBeVisible();
  await expect(page.getByText('Candidate Match Intelligence')).toBeVisible();

  await page.getByRole('button', { name: 'Skills Matrix' }).click();
  await expect(page.getByText('cloud devops', { exact: true })).toBeVisible();

  await page.getByRole('slider').fill('95');
  await page.getByRole('button', { name: 'Run TF-IDF & Cosine Ranking' }).click();
  await page.getByRole('button', { name: 'Candidate Rankings' }).click();
  await expect(page.getByText('No candidates met the minimum threshold.')).toBeVisible();
});
