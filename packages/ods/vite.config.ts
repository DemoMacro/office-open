import { defineConfig } from "vite-plus";

export default defineConfig({
  pack: {
    entry: ["src/index.ts"],
    deps: {
      dts: {
        alwaysBundle: ["@office-open/core"],
      },
    },
    shims: true,
  },
});
