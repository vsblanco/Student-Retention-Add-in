import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    globals: true,
    // node environment — pure-utility tests don't need a DOM.
    // Switch to 'jsdom' if/when we add tests for code that touches `window`.
    environment: 'node',
    include: ['__tests__/**/*.test.js', 'src/**/*.test.js'],
  },
});
