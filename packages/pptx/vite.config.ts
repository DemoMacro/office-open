import path from "node:path";

import { defineConfig } from "vite-plus";

const src = path.resolve("src");

export default defineConfig({
    resolve: {
        alias: {
            "@file": path.resolve(src, "file"),
            "@export": path.resolve(src, "export"),
            "@util": path.resolve(src, "util"),
            tests: path.resolve(src, "tests"),
        },
    },
});
