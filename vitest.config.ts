import { defineConfig } from "vitest/config";

export default defineConfig({
  resolve: {
    alias: {
      hyperformula: "hyperformula/es/index.js",
    },
  },
  test: {
    environment: "node",
  },
});
