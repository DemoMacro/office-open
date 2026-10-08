import { describe, expect, it } from "vite-plus/test";

import { wirePartHyperlinks } from "./shared";

describe("wirePartHyperlinks", () => {
  it("keeps anchor relationships internal and external URLs external", () => {
    const entries: { target: string; mode?: string }[] = [];
    wirePartHyperlinks(
      '<a:rPr r:id="{hlink:anchor}"/><a:rPr r:id="{hlink:external}"/>',
      [
        { key: "anchor", url: "#slide-anchor" },
        { key: "external", url: "https://example.com/" },
      ],
      1,
      (_id, _type, target, mode) => entries.push({ target, mode }),
      "",
    );
    expect(entries).toEqual([
      { target: "#slide-anchor", mode: undefined },
      { target: "https://example.com/", mode: "External" },
    ]);
  });
});
