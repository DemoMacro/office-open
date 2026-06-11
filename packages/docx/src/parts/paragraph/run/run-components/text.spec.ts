import { SpaceType } from "@shared/constants";
import { describe, expect, it } from "vite-plus/test";

import { buildText } from "./text";

describe("buildText", () => {
  it("builds text element with string argument", () => {
    expect(buildText(" this is\n text")).toEqual({
      "w:t": [{ _attr: { "xml:space": "preserve" } }, " this is\n text"],
    });
  });

  it("builds text element with options and explicit space type", () => {
    expect(
      buildText({
        space: SpaceType.PRESERVE,
        text: " this is\n text",
      }),
    ).toEqual({
      "w:t": [{ _attr: { "xml:space": "preserve" } }, " this is\n text"],
    });
  });

  it("builds text element with options and default space type", () => {
    expect(
      buildText({
        text: " this is\n text",
      }),
    ).toEqual({
      "w:t": [{ _attr: { "xml:space": "default" } }, " this is\n text"],
    });
  });
});
