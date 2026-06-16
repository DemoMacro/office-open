import { diffTagSets, roundTripFields, findFieldSpec } from "@office-open/core/descriptor";
import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { describe, it, expect } from "vite-plus/test";

import { corePropertiesDesc } from "./core-properties";
import type { CorePropertiesInput } from "./core-properties";

const spec = findFieldSpec("core-properties")!;

describe("core-properties field consistency", () => {
  it("flags created/modified as F2 write-only inflation and F3 parse-loss", () => {
    // stringify unconditionally emits dcterms:created + dcterms:modified but
    // they have no home on CorePropertiesInput (F2) and parse never reads
    // them back (F3) — the same two fields violate both dimensions.
    const report = diffTagSets(spec);
    expect(report.f2WriteOnly).toEqual(["created", "modified"]);
    expect(report.f3ParseLoss).toEqual(["created", "modified"]);
    expect(report.f1WriteLoss).toEqual([]);
    expect(report.f5ParseOnly).toEqual([]);
  });

  it("round-trips the 8 interface fields without loss or mutation (F4)", () => {
    // Wrap as arrows: corePropertiesDesc.stringify/parse are object methods, so
    // a bare reference trips unbound-method; the descriptors never use `this`.
    // Both ignore ctx, so stub casts are enough at runtime.
    const result = roundTripFields(
      (opts, ctx) => corePropertiesDesc.stringify(opts, ctx),
      (el, ctx) => corePropertiesDesc.parse(el, ctx),
      spec.sampleOptions as CorePropertiesInput,
      {} as unknown as WriteContext,
      {} as unknown as ReadContext,
    );
    expect(result.lost).toEqual([]);
    expect(result.gained).toEqual([]);
    expect(result.mutated).toEqual([]);
  });
});
