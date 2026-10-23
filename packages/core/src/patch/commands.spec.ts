import { describe, expect, it } from "vite-plus/test";

import { applyPatchCommands } from "./commands";
import { PatchError } from "./diagnostics";

describe("canonical patch commands", () => {
  it("rejects unknown command types instead of skipping them", () => {
    const options = { values: ["before"] };
    const command = { type: "replace-all", at: { field: "values" }, value: ["after"] };

    expect(() => applyPatchCommands(options, [command as never])).toThrowError(PatchError);
    expect(options.values).toEqual(["before"]);
  });
});
