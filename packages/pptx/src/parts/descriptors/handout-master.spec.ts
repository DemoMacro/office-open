import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { handoutMasterDesc } from "./handout-master";
import type { HandoutMasterDescriptorOptions } from "./handout-master";

// ── Mock contexts ──

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: HandoutMasterDescriptorOptions) {
  const xml = handoutMasterDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return handoutMasterDesc.parse(el, readCtx);
}

describe("handoutMasterDesc round-trip", () => {
  it("round-trips with default options", () => {
    const opts: HandoutMasterDescriptorOptions = {};
    const result = roundTrip(opts);

    // Default color map values
    expect(result.options).toBeDefined();
    expect(result.options!.colorMap).toBeDefined();
    expect(result.options!.colorMap!.bg1).toBe("lt1");
    expect(result.options!.colorMap!.tx1).toBe("dk1");
    expect(result.options!.colorMap!.bg2).toBe("lt2");
    expect(result.options!.colorMap!.tx2).toBe("dk2");
    expect(result.options!.colorMap!.accent1).toBe("accent1");
    expect(result.options!.colorMap!.accent2).toBe("accent2");
    expect(result.options!.colorMap!.accent3).toBe("accent3");
    expect(result.options!.colorMap!.accent4).toBe("accent4");
    expect(result.options!.colorMap!.accent5).toBe("accent5");
    expect(result.options!.colorMap!.accent6).toBe("accent6");
    expect(result.options!.colorMap!.hlink).toBe("hlink");
    expect(result.options!.colorMap!.folHlink).toBe("folHlink");

    // Default header/footer values
    expect(result.options!.headerFooter).toBeDefined();
    expect(result.options!.headerFooter!.date).toBe(false);
    expect(result.options!.headerFooter!.header).toBe(false);
    expect(result.options!.headerFooter!.footer).toBe(false);
    expect(result.options!.headerFooter!.slideNumber).toBe(false);
  });

  it("round-trips custom color map", () => {
    const opts: HandoutMasterDescriptorOptions = {
      options: {
        colorMap: {
          bg1: "customBg1",
          tx1: "customTx1",
          accent1: "customAccent1",
        },
      },
    };
    const result = roundTrip(opts);

    expect(result.options!.colorMap!.bg1).toBe("customBg1");
    expect(result.options!.colorMap!.tx1).toBe("customTx1");
    expect(result.options!.colorMap!.accent1).toBe("customAccent1");
    // Other values should remain default
    expect(result.options!.colorMap!.bg2).toBe("lt2");
    expect(result.options!.colorMap!.tx2).toBe("dk2");
  });

  it("round-trips header footer settings", () => {
    const opts: HandoutMasterDescriptorOptions = {
      options: {
        headerFooter: {
          date: true,
          header: true,
          footer: false,
          slideNumber: true,
        },
      },
    };
    const result = roundTrip(opts);

    expect(result.options!.headerFooter!.date).toBe(true);
    expect(result.options!.headerFooter!.header).toBe(true);
    expect(result.options!.headerFooter!.footer).toBe(false);
    expect(result.options!.headerFooter!.slideNumber).toBe(true);
  });

  it("round-trips all options together", () => {
    const opts: HandoutMasterDescriptorOptions = {
      options: {
        colorMap: {
          bg1: "white",
          tx1: "black",
          accent1: "blue",
          hlink: "purple",
        },
        headerFooter: {
          date: true,
          header: false,
          footer: true,
          slideNumber: false,
        },
      },
    };
    const result = roundTrip(opts);

    expect(result.options!.colorMap!.bg1).toBe("white");
    expect(result.options!.colorMap!.tx1).toBe("black");
    expect(result.options!.colorMap!.accent1).toBe("blue");
    expect(result.options!.colorMap!.hlink).toBe("purple");

    expect(result.options!.headerFooter!.date).toBe(true);
    expect(result.options!.headerFooter!.header).toBe(false);
    expect(result.options!.headerFooter!.footer).toBe(true);
    expect(result.options!.headerFooter!.slideNumber).toBe(false);
  });
});
