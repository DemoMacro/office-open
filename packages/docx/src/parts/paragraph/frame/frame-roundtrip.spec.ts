import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { parseParagraphProperties } from "../../../body";
import type { DocxReadContext } from "../../../context";
import { stringifyParagraphProperties } from "../stringify";

const W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

const readCtx = {} as unknown as DocxReadContext;

function roundTripFrame(frame: Record<string, unknown>): Record<string, unknown> {
  const { xml } = stringifyParagraphProperties({ frame: frame as never });
  // xml is a complete <w:pPr>...</w:pPr>; declare the w: namespace on it
  const doc = parseXml(xml!.replace("<w:pPr>", `<w:pPr ${W_NS}>`));
  const pPr = doc.elements![0];
  const opts = parseParagraphProperties(pPr, readCtx);
  return opts.frame as Record<string, unknown>;
}

describe("framePr parse round-trip", () => {
  it("round-trips alignment.x/y (xAlign/yAlign)", () => {
    const result = roundTripFrame({
      type: "alignment",
      alignment: { x: "center", y: "top" },
      anchor: { horizontal: "text", vertical: "text" },
      width: 1000,
      height: 500,
    });
    expect(result.alignment).toBeDefined();
    expect((result.alignment as Record<string, unknown>).x).toBe("center");
    expect((result.alignment as Record<string, unknown>).y).toBe("top");
  });

  it("round-trips anchor.horizontal/vertical (hAnchor/vAnchor)", () => {
    const result = roundTripFrame({
      type: "alignment",
      alignment: { x: "center", y: "top" },
      anchor: { horizontal: "page", vertical: "margin" },
      width: 1000,
      height: 500,
    });
    expect(result.anchor).toBeDefined();
    expect((result.anchor as Record<string, unknown>).horizontal).toBe("page");
    expect((result.anchor as Record<string, unknown>).vertical).toBe("margin");
  });

  it("round-trips anchorLock", () => {
    const result = roundTripFrame({
      type: "alignment",
      alignment: { x: "center", y: "top" },
      anchor: { horizontal: "text", vertical: "text" },
      anchorLock: true,
      width: 1000,
      height: 500,
    });
    expect(result.anchorLock).toBe(true);
  });
});
