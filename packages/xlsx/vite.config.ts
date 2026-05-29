import nodePolyfills from "@rolldown/plugin-node-polyfills";
import { defineConfig } from "vite-plus";

export default defineConfig({
  pack: {
    entry: ["src/index.ts"],
    plugins: [nodePolyfills()],
    platform: "neutral",
    shims: true,
  },
});
