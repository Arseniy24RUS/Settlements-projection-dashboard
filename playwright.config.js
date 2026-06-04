const { defineConfig, devices } = require('@playwright/test');

const PORT = 4175;
const baseURL = `http://127.0.0.1:${PORT}`;

module.exports = defineConfig({
  testDir: './tests',
  timeout: 180000,
  expect: { timeout: 30000 },
  fullyParallel: false,
  workers: 1,
  reporter: [['list']],
  use: {
    baseURL,
    ignoreHTTPSErrors: true,
    trace: 'retain-on-failure'
  },
  webServer: {
    command: `npx http-server . -a 127.0.0.1 -p ${PORT} -c-1 --silent`,
    url: baseURL,
    reuseExistingServer: true,
    timeout: 120000
  },
  projects: [
    {
      name: 'desktop-en',
      use: {
        locale: 'en-US',
        viewport: { width: 1440, height: 1100 }
      }
    },
    {
      name: 'desktop-ru',
      use: {
        locale: 'ru-RU',
        viewport: { width: 1440, height: 1100 }
      }
    },
    {
      name: 'mobile-en',
      use: {
        ...devices['Pixel 5'],
        locale: 'en-US'
      }
    },
    {
      name: 'mobile-ru',
      use: {
        ...devices['Pixel 5'],
        locale: 'ru-RU'
      }
    }
  ]
});
