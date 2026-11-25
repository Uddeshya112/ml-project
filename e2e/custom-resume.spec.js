const { test, expect } = require('@playwright/test');

test('custom resume can be added and shortlisted', async ({ page }) => {
  await page.goto('/', { waitUntil: 'domcontentloaded' });

  await page.locator('#openAddResumeModalBtn').click();
  await page.getByLabel('Candidate Identifier').fill('Test Cloud Engineer');
  await page.getByLabel('Resume Content').fill('Python AWS Docker Kubernetes Terraform Linux CI/CD');
  await page.getByRole('button', { name: 'Add to Corpus' }).click();

  await expect(page.getByText('Test Cloud Engineer', { exact: true })).toBeVisible();
  await page.getByRole('button', { name: 'Run TF-IDF & Cosine Ranking' }).click();
  await expect(page.getByText('Candidate Match Intelligence')).toBeVisible();

  const customPanel = page.getByRole('heading', { name: 'Test Cloud Engineer' })
    .locator("xpath=ancestor::div[contains(concat(' ', normalize-space(@class), ' '), ' panel ')][1]");
  await customPanel.getByRole('button', { name: '+ Shortlist' }).click();
  await expect(page.getByRole('button', { name: 'Shortlisted ✓' }).first()).toBeVisible();
  await page.locator('#open-bag-btn').click();
  await expect(page.locator('#bag-items').getByText('Test Cloud Engineer', { exact: true })).toBeVisible();
});
