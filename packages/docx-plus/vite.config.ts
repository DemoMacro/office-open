import { defineConfig } from "vite-plus";

export default defineConfig({
  pack: {
    entry: ["src/index.ts"],
    format: ["esm", "cjs"],
    platform: "neutral",
    shims: true,
  },
});
