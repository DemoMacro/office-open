import { configDefaults, defineConfig } from "vite-plus";

export default defineConfig({
    fmt: {
        sortImports: {
            type: "natural",
        },
        sortPackageJson: true,
        sortTailwindcss: {},
    },
    lint: {
        ignorePatterns: ["**/scripts/**", "**/dist/**", "**/coverage/**"],
        options: {
            typeAware: true,
            typeCheck: true,
        },
        rules: {
            "no-redundant-type-constituents": "off",
            "no-unused-expressions": "off",
        },
    },
    staged: {
        "*": "vp check --fix",
    },
    test: {
        coverage: {
            exclude: [
                ...configDefaults.exclude,
                "**/dist/**",
                "**/demo/**",
                "**/docs/**",
                "**/scripts/**",
                "**/src/**/index.ts",
                "**/src/**/types.ts",
                "**/src/util/output-type.ts",
                "**/*.spec.ts",
            ],
            provider: "v8",
            reporter: ["text", "json", "html"],
            thresholds: {
                branches: 95,
                functions: 99,
                lines: 98,
                statements: 98,
            },
        },
        environment: "happy-dom",
        exclude: [
            ...configDefaults.exclude,
            "**/build/**",
            "**/demo/**",
            "**/docs/**",
            "**/scripts/**",
        ],
        include: ["**/src/**/*.spec.ts", "**/packages/**/*.spec.ts"],
    },
});
