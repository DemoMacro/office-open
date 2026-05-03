import { defineBuildConfig } from "@funish/basis/config";
import type { BuildConfig } from "@funish/build/config";
import nodePolyfills from "@rolldown/plugin-node-polyfills";

export default defineBuildConfig({
    entries: [
        {
            entry: ["src/index.ts"],
            plugins: [nodePolyfills()],
            shims: true,
            deps: {
                neverBundle: ["stream"],
            },
        },
    ],
} satisfies BuildConfig);
