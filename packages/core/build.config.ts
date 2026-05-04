import { defineBuildConfig } from "@funish/basis/config";
import type { BuildConfig } from "@funish/build/config";

export default defineBuildConfig({
    entries: [
        {
            entry: [
                "src/index.ts",
                "src/values.ts",
                "src/smartart/index.ts",
                "src/chart/index.ts",
                "src/drawingml/index.ts",
                "src/archive.ts",
            ],
        },
    ],
} satisfies BuildConfig);
