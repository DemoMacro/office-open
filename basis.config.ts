import { defineBasisConfig } from "@funish/basis/config";

export default defineBasisConfig({
    publish: {
        npm: {
            additionalTag: "edge",
        },
    },
    git: {
        hooks: {
            "pre-commit": "pnpm basis git staged",
            "commit-msg": "pnpm basis git lint-commit",
        },
        staged: {
            rules: {
                "*": "pnpm check",
            },
        },
    },
});
