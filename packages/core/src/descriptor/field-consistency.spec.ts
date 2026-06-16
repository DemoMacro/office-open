import type { Element } from "@office-open/xml";
import { describe, it, expect } from "vite-plus/test";

import { checkOrder, diffTagSets, roundTripFields } from "./field-consistency";
import type { DescriptorFieldSpec } from "./field-spec";

const spec = (over: Partial<DescriptorFieldSpec>): DescriptorFieldSpec => ({
  id: "test",
  optionsInterface: "Test",
  interfaceFields: ["a", "b"],
  writeFields: ["a", "b"],
  parseFields: ["a", "b"],
  sampleOptions: {},
  ...over,
});

describe("diffTagSets", () => {
  it("returns no drift for a symmetric spec", () => {
    const r = diffTagSets(spec({}));
    expect(r.f1WriteLoss).toEqual([]);
    expect(r.f2WriteOnly).toEqual([]);
    expect(r.f3ParseLoss).toEqual([]);
    expect(r.f5ParseOnly).toEqual([]);
  });

  it("F1 — flags interface fields never written", () => {
    expect(diffTagSets(spec({ writeFields: ["a"] })).f1WriteLoss).toEqual(["b"]);
  });

  it("F2 — flags write-only inflation not on the interface", () => {
    expect(diffTagSets(spec({ writeFields: ["a", "b", "c"] })).f2WriteOnly).toEqual(["c"]);
  });

  it("F3 — flags written fields never parsed back", () => {
    expect(diffTagSets(spec({ parseFields: ["a"] })).f3ParseLoss).toEqual(["b"]);
  });

  it("F5 — flags parse-only fields never written", () => {
    expect(diffTagSets(spec({ writeFields: ["a"], parseFields: ["a", "b"] })).f5ParseOnly).toEqual([
      "b",
    ]);
  });
});

describe("roundTripFields", () => {
  const stringify = (opts: Record<string, unknown>): string =>
    `<root>${Object.entries(opts)
      .filter(([, v]) => v !== undefined)
      .map(([k, v]) => `<${k}>${String(v)}</${k}>`)
      .join("")}</root>`;
  const parse = (el: Element): Record<string, unknown> => {
    const result: Record<string, unknown> = {};
    for (const child of el.elements ?? []) {
      if (child.type === "element" && child.name) result[child.name] = child.elements?.[0]?.text;
    }
    return result;
  };

  it("returns no drift when parse restores every field", () => {
    const r = roundTripFields(stringify, parse, { a: "1", b: "2" }, {}, {});
    expect(r.lost).toEqual([]);
    expect(r.gained).toEqual([]);
    expect(r.mutated).toEqual([]);
  });

  it("reports lost fields when parse drops one", () => {
    const drop = (el: Element): Record<string, unknown> => {
      const result: Record<string, unknown> = {};
      for (const child of el.elements ?? []) {
        if (child.type === "element" && child.name && child.name !== "b") {
          result[child.name] = child.elements?.[0]?.text;
        }
      }
      return result;
    };
    expect(roundTripFields(stringify, drop, { a: "1", b: "2" }, {}, {}).lost).toEqual(["b"]);
  });

  it("reports gained fields when parse invents one", () => {
    const invent = (el: Element): Record<string, unknown> => ({ ...parse(el), extra: "x" });
    expect(roundTripFields(stringify, invent, { a: "1" }, {}, {}).gained).toEqual(["extra"]);
  });

  it("reports mutated fields when a value changes", () => {
    const mutate = (el: Element): Record<string, unknown> => ({ ...parse(el), a: "changed" });
    expect(roundTripFields(stringify, mutate, { a: "1" }, {}, {}).mutated).toEqual(["a"]);
  });
});

describe("checkOrder", () => {
  const order = ["a", "b", "c", "d"];

  it("returns no violations when children follow the sequence", () => {
    expect(checkOrder(`<root><a/><b/><c/></root>`, order)).toEqual([]);
  });

  it("flags a child that precedes an earlier sibling in the sequence", () => {
    // a(0) c(2) b(1) — b lands before c, violating the order.
    const v = checkOrder(`<root><a/><c/><b/></root>`, order);
    expect(v).toEqual([{ index: 2, tag: "b", reason: "out-of-order" }]);
  });

  it("flags children absent from the expected sequence", () => {
    const v = checkOrder(`<root><a/><x/><b/></root>`, order);
    expect(v).toEqual([{ index: 1, tag: "x", reason: "unexpected" }]);
  });

  it("normalizes namespace prefixes on both sides", () => {
    const xml = `<root xmlns:w="u"><w:a/><w:b/></root>`;
    expect(checkOrder(xml, ["a", "b"])).toEqual([]);
  });
});
