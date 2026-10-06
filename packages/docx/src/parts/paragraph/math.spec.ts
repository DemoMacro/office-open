import { parse } from "@office-open/xml";
import { describe, expect, it } from "vitest";

import { parseMathChildren, stringifyMathInput } from "./math/stringify";

describe("math run text", () => {
  it("round-trips preserved spaces", () => {
    const xml =
      '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">' +
      '<m:r><m:t xml:space="preserve"> x </m:t></m:r></m:oMath>';
    const root = parse(xml).elements?.[0];
    if (!root) throw new Error("math XML did not parse");
    const children = parseMathChildren(root);
    expect(children).toEqual([{ text: " x ", preserveSpace: true }]);
    expect(stringifyMathInput(children[0]!)).toContain('<m:t xml:space="preserve"> x </m:t>');
  });
});
