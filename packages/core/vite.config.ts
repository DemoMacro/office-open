import { defineConfig } from "vite-plus";

export default defineConfig({
  pack: {
    entry: [
      "src/index.ts",
      "src/values.ts",
      "src/smartart/index.ts",
      "src/chart/index.ts",
      "src/drawingml/index.ts",
      "src/patch/index.ts",
    ],
    shims: true,
  },
  test: {
    benchmark: {
      include: ["src/**/*.bench.ts"],
    },
  },
});
