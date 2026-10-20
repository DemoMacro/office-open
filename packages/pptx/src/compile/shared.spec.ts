import { Relationships, RELATIONSHIP_TYPES } from "@office-open/core/opc";
import { describe, expect, it } from "vite-plus/test";

import { claimMediaSourceRelationships, wirePartHyperlinks } from "./shared";

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

describe("claimMediaSourceRelationships", () => {
  it("preserves duplicate source media references used by markup branches", () => {
    const rels = new Relationships();
    rels.addRelationship(1, RELATIONSHIP_TYPES.slideLayout, "../slideLayouts/slideLayout.xml");
    rels.addRelationship(2, RELATIONSHIP_TYPES.audio, "../media/audio.wav");
    claimMediaSourceRelationships(rels, "ppt/slides/slide1.xml", [
      {
        source: "ppt/slides/slide1.xml",
        relationshipType: RELATIONSHIP_TYPES.audio,
        target: "../media/audio.wav",
        rId: "rId2",
      },
      {
        source: "ppt/slides/slide1.xml",
        relationshipType: RELATIONSHIP_TYPES.audio,
        target: "../media/audio.wav",
        rId: "rId3",
      },
    ]);

    const xml = rels.serialize();
    expect(xml).toContain('Id="rId2"');
    expect(xml).toContain('Id="rId3"');
  });
});
