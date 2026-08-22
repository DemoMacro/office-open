import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { WebSettingsOptions } from "./web-settings";
import { webSettingsDesc } from "./web-settings";

const writeCtx = {} as unknown as WriteContext;
const readCtx = {} as unknown as ReadContext;

function roundTrip(opts: WebSettingsOptions) {
  const xml = webSettingsDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return webSettingsDesc.parse(el, readCtx) as unknown as WebSettingsOptions;
}

describe("webSettingsDesc round-trip", () => {
  it("round-trips optimizeForBrowser", () => {
    const result = roundTrip({ optimizeForBrowser: true });
    expect(result.optimizeForBrowser).toBe(true);
  });

  it("round-trips optimizeForBrowser with target", () => {
    const result = roundTrip({ optimizeForBrowser: { value: true, target: "1024x768" } });
    expect(result.optimizeForBrowser).toEqual({ value: true, target: "1024x768" });
  });

  it("round-trips optimizeForBrowser disabled with target", () => {
    const result = roundTrip({ optimizeForBrowser: { value: false, target: "800x600" } });
    expect(result.optimizeForBrowser).toEqual({ value: false, target: "800x600" });
  });

  it("round-trips allowPNG", () => {
    const result = roundTrip({ allowPNG: true });
    expect(result.allowPNG).toBe(true);
  });

  it("round-trips pixelsPerInch", () => {
    const result = roundTrip({ pixelsPerInch: 96 });
    expect(result.pixelsPerInch).toBe(96);
  });

  it("round-trips encoding", () => {
    const result = roundTrip({ encoding: "utf-8" });
    expect(result.encoding).toBe("utf-8");
  });

  it("round-trips targetScreenSize", () => {
    const result = roundTrip({ targetScreenSize: "1024x768" });
    expect(result.targetScreenSize).toBe("1024x768");
  });

  it("round-trips divs with margins and borders", () => {
    const result = roundTrip({
      divs: [
        {
          id: 100,
          marginLeft: 720,
          marginRight: 720,
          marginTop: 360,
          marginBottom: 360,
          blockQuote: true,
        },
      ],
    });
    expect(result.divs).toHaveLength(1);
    expect(result.divs![0]?.id).toBe(100);
    expect(result.divs![0]?.marginLeft).toBe(720);
    expect(result.divs![0]?.blockQuote).toBe(true);
  });

  it("round-trips empty web settings", () => {
    const result = roundTrip({});
    expect(Object.keys(result).length).toBe(0);
  });

  it("round-trips div borders with an ST_Border style", () => {
    const result = roundTrip({
      divs: [
        {
          id: 1,
          marginLeft: 720,
          marginRight: 720,
          marginTop: 360,
          marginBottom: 360,
          border: { top: { style: "single", color: "FF0000", size: 4 } },
        },
      ],
    });
    expect(result.divs![0]?.border?.top).toEqual({ style: "single", color: "FF0000", size: 4 });
  });

  it("parses an unknown ST_Border token through verbatim", () => {
    const result = roundTrip({
      divs: [
        {
          id: 1,
          marginLeft: 0,
          marginRight: 0,
          marginTop: 0,
          marginBottom: 0,
          border: { left: { style: "single" } },
        },
      ],
    });
    expect(result.divs![0]?.border?.left?.style).toBe("single");
  });

  it("throws when a div border side carries no style", () => {
    expect(() =>
      webSettingsDesc.stringify(
        {
          divs: [
            {
              id: 1,
              marginLeft: 0,
              marginRight: 0,
              marginTop: 0,
              marginBottom: 0,
              border: { top: {} as never },
            },
          ],
        },
        writeCtx,
      ),
    ).toThrow(/div border.*style/);
  });

  it("throws when a div misses the required id or margins", () => {
    expect(() =>
      webSettingsDesc.stringify(
        { divs: [{ marginLeft: 0, marginRight: 0, marginTop: 0, marginBottom: 0 } as never] },
        writeCtx,
      ),
    ).toThrow(/CT_Div/);
  });

  it("round-trips combined options", () => {
    const result = roundTrip({
      optimizeForBrowser: true,
      allowPNG: false,
      pixelsPerInch: 120,
      doNotRelyOnCSS: true,
    });
    expect(result.optimizeForBrowser).toBe(true);
    expect(result.allowPNG).toBe(false);
    expect(result.pixelsPerInch).toBe(120);
    expect(result.doNotRelyOnCSS).toBe(true);
  });
});
