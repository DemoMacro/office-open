import { defineBuildConfig } from "@funish/basis/config";
import type { BuildConfig } from "@funish/build/config";

export default defineBuildConfig({
    entries: [
        {
            entry: ["src/index", "src/values"],
            deps: {
                alwaysBundle: ["@office-open/xml"],
                onlyBundle: false,
            },
        },
    ],
} satisfies BuildConfig);
