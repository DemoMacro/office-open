import type { CorePropertiesOptions } from "@office-open/core";
import { diffTagSets, roundTripFields, findFieldSpec } from "@office-open/core/descriptor";
import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { describe, it, expect } from "vite-plus/test";

import { corePropertiesDesc } from "./core-properties";

const spec = findFieldSpec("core-properties")!;

describe("core-properties field consistency", () => {
  it("round-trips created/modified with no F2 inflation or F3 parse-loss", () => {
    // created/modified now have a home on CorePropertiesOptions and parse reads
    // them back, so the field sets are fully symmetric.
    const report = diffTagSets(spec);
    expect(report.f1WriteLoss).toEqual([]);
    expect(report.f2WriteOnly).toEqual([]);
    expect(report.f3ParseLoss).toEqual([]);
    expect(report.f5ParseOnly).toEqual([]);
  });

  it("round-trips the 8 interface fields without loss or mutation (F4)", () => {
    // Wrap as arrows: corePropertiesDesc.stringify/parse are object methods, so
    // a bare reference trips unbound-method; the descriptors never use `this`.
    // Both ignore ctx, so stub casts are enough at runtime.
    const result = roundTripFields(
      (opts, ctx) => corePropertiesDesc.stringify(opts, ctx),
      (el, ctx) => corePropertiesDesc.parse(el, ctx),
      spec.sampleOptions as CorePropertiesOptions,
      {} as unknown as WriteContext,
      {} as unknown as ReadContext,
    );
    expect(result.lost).toEqual([]);
    expect(result.gained).toEqual([]);
    expect(result.mutated).toEqual([]);
  });
});
