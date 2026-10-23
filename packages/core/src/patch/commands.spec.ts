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

  it("replaces only the first text match unless all is selected", () => {
    const options = { items: [{ children: ["a-a"] }] };

    applyPatchCommands(options, [
      {
        type: "replace",
        at: { field: "items", route: [{ index: 0 }], text: { literal: "a" } },
        value: "b",
      },
    ]);
    expect(options.items[0]?.children).toEqual(["b-a"]);

    applyPatchCommands(options, [
      {
        type: "replace",
        at: { field: "items", route: [{ index: 0 }], text: { literal: "a", all: true } },
        value: "c",
      },
    ]);
    expect(options.items[0]?.children).toEqual(["b-c"]);
  });
});
