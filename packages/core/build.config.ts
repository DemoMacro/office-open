import { defineBuildConfig } from "@funish/basis/config";
import type { BuildConfig } from "@funish/build/config";

export default defineBuildConfig({
    entries: [
        {
            entry: [
                "src/index",
                "src/values",
                "src/smartart/index",
                "src/smartart/categories",
                "src/smartart/smartart-collection",
                "src/smartart/tree-to-model",
                "src/smartart/data-model/index",
                "src/smartart/data-model/connection",
                "src/smartart/data-model/data-model",
                "src/smartart/data-model/point",
            ],
            deps: {
                alwaysBundle: ["@office-open/xml"],
                onlyBundle: false,
            },
        },
    ],
} satisfies BuildConfig);
