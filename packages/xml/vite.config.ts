import { defineConfig } from "vite-plus";

export default defineConfig({
  pack: {
    entry: ["src/index.ts", "src/utils.ts"],
    shims: true,
  },
  test: {
    benchmark: {
      include: ["src/**/*.bench.ts"],
    },
  },
});
