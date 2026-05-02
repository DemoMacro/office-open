import { defineBuildConfig } from "@funish/basis/config";
import type { BuildConfig } from "@funish/build/config";
import nodePolyfills from "@rolldown/plugin-node-polyfills";

export default defineBuildConfig({
    entries: [
        {
            entry: ["src/index"],
            plugins: [nodePolyfills()],
            shims: true,
            external: ["stream"],
            deps: {
                alwaysBundle: ["@office-open/xml", "@office-open/core"],
                onlyBundle: false,
            },
        },
    ],
} satisfies BuildConfig);
