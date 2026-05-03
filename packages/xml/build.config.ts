import { defineBuildConfig } from "@funish/basis/config";
import type { BuildConfig } from "@funish/build/config";

export default defineBuildConfig({
    entries: [
        {
            entry: ["src/index.ts"],
        },
    ],
} satisfies BuildConfig);
