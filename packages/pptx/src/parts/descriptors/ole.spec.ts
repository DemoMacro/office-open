import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { oleDesc } from "./ole";
import type { OleDescriptorOptions } from "./ole";

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

function roundTrip(opts: OleDescriptorOptions) {
  const xml = oleDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements![0];
  return oleDesc.parse(el, readCtx);
}

describe("oleDesc round-trip", () => {
  it("round-trips basic embedded OLE object", () => {
    const opts: OleDescriptorOptions = {
      id: 100,
      name: "Test OLE",
      x: 50,
      y: 60,
      width: 200,
      height: 150,
      progId: "Excel.Sheet.12",
      embed: { rId: "rId2" },
    };
    const result = roundTrip(opts);

    expect(result.id).toBe(100);
    expect(result.name).toBe("Test OLE");
    expect(result.x).toBe(50);
    expect(result.y).toBe(60);
    expect(result.width).toBe(200);
    expect(result.height).toBe(150);
    expect(result.progId).toBe("Excel.Sheet.12");
    expect(result.embed).toBeDefined();
    expect(result.embed!.rId).toBe("rId2");
  });

  it("round-trips linked OLE object", () => {
    const opts: OleDescriptorOptions = {
      id: 200,
      name: "Linked OLE",
      x: 10,
      y: 20,
      width: 300,
      height: 200,
      progId: "Word.Document.12",
      link: { rId: "rId3", autoUpdate: true },
    };
    const result = roundTrip(opts);

    expect(result.id).toBe(200);
    expect(result.name).toBe("Linked OLE");
    expect(result.progId).toBe("Word.Document.12");
    expect(result.link).toBeDefined();
    expect(result.link!.rId).toBe("rId3");
    expect(result.link!.autoUpdate).toBe(true);
  });

  it("round-trips linked OLE without autoUpdate", () => {
    const opts: OleDescriptorOptions = {
      id: 250,
      link: { rId: "rId4" },
    };
    const result = roundTrip(opts);

    expect(result.link).toBeDefined();
    expect(result.link!.rId).toBe("rId4");
    expect(result.link!.autoUpdate).toBeUndefined();
  });

  it("round-trips OLE with showAsIcon", () => {
    const opts: OleDescriptorOptions = {
      id: 300,
      name: "Icon OLE",
      showAsIcon: true,
      imgW: 64,
      imgH: 64,
      embed: { rId: "rId5" },
    };
    const result = roundTrip(opts);

    expect(result.id).toBe(300);
    expect(result.showAsIcon).toBe(true);
    expect(result.imgW).toBe(64);
    expect(result.imgH).toBe(64);
  });

  it("round-trips OLE with spid", () => {
    const opts: OleDescriptorOptions = {
      id: 400,
      spid: "_x0000_s1025",
      embed: { rId: "rId6" },
    };
    const result = roundTrip(opts);

    expect(result.spid).toBe("_x0000_s1025");
  });

  it("round-trips OLE with followColorScheme", () => {
    const opts: OleDescriptorOptions = {
      id: 500,
      followColorScheme: "full",
      embed: { rId: "rId7" },
    };
    const result = roundTrip(opts);

    expect(result.followColorScheme).toBe("full");
  });

  it("round-trips position with defaults", () => {
    const opts: OleDescriptorOptions = {
      id: 700,
      embed: { rId: "rId10" },
    };
    const result = roundTrip(opts);

    expect(result.id).toBe(700);
    expect(result.name).toBe("Object 700");
    expect(result.x).toBe(0);
    expect(result.y).toBe(0);
    // default 100px = 952500 EMU
    expect(result.width).toBe(952500);
    expect(result.height).toBe(952500);
  });

  it("round-trips EMU conversion correctly", () => {
    const opts: OleDescriptorOptions = {
      id: 800,
      x: 1024,
      y: 768,
      width: 1920,
      height: 1080,
      embed: { rId: "rId11" },
    };
    const result = roundTrip(opts);

    expect(result.x).toBe(1024);
    expect(result.y).toBe(768);
    expect(result.width).toBe(1920);
    expect(result.height).toBe(1080);
  });
});
