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

  it("tracks repeated string leaves by their actual array indexes", () => {
    const options = { values: ["a", "a"] };

    applyPatchCommands(options, [
      {
        type: "replace",
        at: { field: "values", text: { literal: "aa" } },
        value: "X",
      },
    ]);

    expect(options.values).toEqual(["X", ""]);
  });

  it("replaces across array items without replacing individual items again", () => {
    const options = { items: ["a"] };

    applyPatchCommands(options, [
      {
        type: "replace",
        at: { field: "items", text: { literal: "a", all: true } },
        value: "aa",
      },
    ]);

    expect(options.items).toEqual(["aa"]);
  });

  it("still replaces text in deeply nested array items", () => {
    const options = { items: [{ body: { text: "Hello {{name}}" } }] };

    applyPatchCommands(options, [
      {
        type: "replace",
        at: { field: "items", text: { placeholder: "name" } },
        value: "Alice",
      },
    ]);

    expect(options.items[0]).toEqual({ body: { text: "Hello Alice" } });
  });

  it("rejects empty text selectors", () => {
    const options = { values: ["a"] };

    expect(() =>
      applyPatchCommands(options, [
        {
          type: "replace",
          at: { field: "values", text: { literal: "", all: true } },
          value: "b",
        },
      ]),
    ).toThrowError(PatchError);
    expect(options.values).toEqual(["a"]);
  });

  it("replaces other fields in an item after a cross-leaf match", () => {
    const options = { items: [{ text: "a", note: "a" }] };

    applyPatchCommands(options, [
      {
        type: "replace",
        at: { field: "items", text: { literal: "a", all: true } },
        value: "aa",
      },
    ]);

    expect(options.items[0]).toEqual({ text: "aa", note: "aa" });
  });
});
