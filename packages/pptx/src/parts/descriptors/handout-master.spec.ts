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
  const xml = handoutMasterDesc.stringify(opts, writeCtx as never)!;
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return handoutMasterDesc.parse(el, readCtx);
}

describe("handoutMasterDesc round-trip", () => {
  it("round-trips with default options", () => {
    const opts: HandoutMasterDescriptorOptions = {};
    const result = roundTrip(opts);

    // Default color map values
    expect(result.options).toBeDefined();
    expect(result.options!.colorMapping).toBeDefined();
    expect(result.options!.colorMapping!.background1).toBe("light1");
    expect(result.options!.colorMapping!.text1).toBe("dark1");
    expect(result.options!.colorMapping!.background2).toBe("light2");
    expect(result.options!.colorMapping!.text2).toBe("dark2");
    expect(result.options!.colorMapping!.accent1).toBe("accent1");
    expect(result.options!.colorMapping!.accent2).toBe("accent2");
    expect(result.options!.colorMapping!.accent3).toBe("accent3");
    expect(result.options!.colorMapping!.accent4).toBe("accent4");
    expect(result.options!.colorMapping!.accent5).toBe("accent5");
    expect(result.options!.colorMapping!.accent6).toBe("accent6");
    expect(result.options!.colorMapping!.hyperlink).toBe("hyperlink");
    expect(result.options!.colorMapping!.followedHyperlink).toBe("followedHyperlink");

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
        colorMapping: {
          background1: "dark1",
          text1: "light1",
          accent1: "accent2",
        },
      },
    };
    const result = roundTrip(opts);

    expect(result.options!.colorMapping!.background1).toBe("dark1");
    expect(result.options!.colorMapping!.text1).toBe("light1");
    expect(result.options!.colorMapping!.accent1).toBe("accent2");
    // Other values should remain default
    expect(result.options!.colorMapping!.background2).toBe("light2");
    expect(result.options!.colorMapping!.text2).toBe("dark2");
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
        colorMapping: {
          background1: "light2",
          text1: "dark2",
          accent1: "accent3",
          hyperlink: "followedHyperlink",
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

    expect(result.options!.colorMapping!.background1).toBe("light2");
    expect(result.options!.colorMapping!.text1).toBe("dark2");
    expect(result.options!.colorMapping!.accent1).toBe("accent3");
    expect(result.options!.colorMapping!.hyperlink).toBe("followedHyperlink");

    expect(result.options!.headerFooter!.date).toBe(true);
    expect(result.options!.headerFooter!.header).toBe(false);
    expect(result.options!.headerFooter!.footer).toBe(true);
    expect(result.options!.headerFooter!.slideNumber).toBe(false);
  });
});
