const { test, expect } = require('@playwright/test');
const path = require('node:path');

const screenshotDir = path.join(process.cwd(), 'qa-screenshots', 'Settlements-projection-dashboard');

function expectedLanguage(projectName) {
  return projectName.endsWith('-ru') ? 'ru' : 'en';
}

function otherLanguage(language) {
  return language === 'ru' ? 'en' : 'ru';
}

async function assertLanguage(page, language) {
  await expect(page.locator('html')).toHaveAttribute('lang', language);
  await expect(page.getByTestId('language-toggle')).toHaveText(language === 'ru' ? 'EN' : 'RU');

  if (language === 'ru') {
    await expect(page.getByRole('heading', { name: 'Населённые пункты России: прогноз численности и половозрастного состава населения' })).toBeVisible();
    await expect(page.getByRole('button', { name: 'Применить' })).toBeVisible();
    await expect(page.getByText('Параметры анализа')).toBeVisible();
    await expect(page.getByText('Населённые пункты в выборке')).toBeVisible();
    await expect(page.getByText('Картограмма населённых пунктов')).toBeVisible();
    await expect(page.getByText('С учётом миграции').first()).toBeVisible();
  } else {
    await expect(page.getByRole('heading', { name: 'Russian settlements: population and age-sex forecast' })).toBeVisible();
    await expect(page.getByRole('button', { name: 'Apply' })).toBeVisible();
    await expect(page.getByText('Analysis parameters')).toBeVisible();
    await expect(page.getByText('Settlements in selection')).toBeVisible();
    await expect(page.getByText('Settlement map')).toBeVisible();
    await expect(page.getByText('With migration').first()).toBeVisible();
  }
}

async function assertMethodologyDialog(page, language) {
  const openName = language === 'ru' ? 'Методика' : 'Methodology';
  const heading = language === 'ru' ? 'Описание методики исследования' : 'Research methodology';
  const closeName = language === 'ru' ? 'Закрыть' : 'Close';

  await page.getByRole('button', { name: openName }).click();
  await expect(page.getByRole('heading', { name: heading })).toBeVisible();
  await page.getByRole('button', { name: closeName }).click();
  await expect(page.getByRole('heading', { name: heading })).not.toBeVisible();
}

async function waitForDashboard(page) {
  await page.goto('/');
  await page.waitForFunction(() => window.AppI18n && window.__SETTLEMENTS_APP_READY__ === true, null, { timeout: 180000 });
  await expect(page.locator('#loadingOverlay')).not.toHaveClass(/visible/);
}

test('defaults from browser locale, toggles EN/RU, and captures QA screenshot', async ({ page }, testInfo) => {
  const consoleErrors = [];
  page.on('console', (message) => {
    if (message.type() === 'error') consoleErrors.push(message.text());
  });
  page.on('pageerror', (error) => {
    consoleErrors.push(error.message);
  });

  await waitForDashboard(page);

  const initialLanguage = expectedLanguage(testInfo.project.name);
  await assertLanguage(page, initialLanguage);
  await assertMethodologyDialog(page, initialLanguage);

  await page.screenshot({
    path: path.join(screenshotDir, `${testInfo.project.name}.png`),
    fullPage: true
  });

  await page.getByTestId('language-toggle').click();
  await assertLanguage(page, otherLanguage(initialLanguage));
  await assertMethodologyDialog(page, otherLanguage(initialLanguage));

  await page.getByTestId('language-toggle').click();
  await assertLanguage(page, initialLanguage);

  await expect(page.locator('#errorBanner')).toHaveClass(/hidden/);
  expect(consoleErrors).toEqual([]);
});
