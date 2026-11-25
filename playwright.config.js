const { defineConfig } = require('@playwright/test');

module.exports = defineConfig({
  testDir: './e2e',
  timeout: 30_000,
  workers: 1,
  use: {
    baseURL: 'http://127.0.0.1:5000',
    headless: true,
  },
  reporter: [['list']],
});
