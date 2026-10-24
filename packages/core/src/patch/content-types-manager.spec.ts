import type { Element } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { appendContentType, appendOverride, removeOverride } from "./content-types-manager";
import { PatchError } from "./diagnostics";

const emptyRoot = (): Element => ({ type: "element", name: "Package" });

describe("content types manager", () => {
  it("rejects malformed packages without a Types root", () => {
    expect(() => appendContentType(emptyRoot(), "image/png", "png")).toThrow(PatchError);
    expect(() => appendOverride(emptyRoot(), "/part.xml", "application/xml")).toThrow(PatchError);
    expect(() => removeOverride(emptyRoot(), "/part.xml")).toThrow(PatchError);
  });
});
