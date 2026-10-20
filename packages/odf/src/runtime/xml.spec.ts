import type { Element } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { descendantElements } from "./xml";

describe("ODF XML helpers", () => {
  it("traverses elements with extremely wide fan-out", () => {
    const children = Array.from({ length: 200_000 }, (_, index): Element => ({
      type: "element",
      name: `child-${index}`,
    }));
    const root: Element = { type: "element", name: "root", elements: children };

    expect(descendantElements(root)).toHaveLength(children.length);
  });
});
