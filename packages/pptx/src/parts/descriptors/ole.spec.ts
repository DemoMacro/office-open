import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { OleOptions } from "../ole-frame";
import { oleDesc } from "./ole";

// ── Mock contexts ──

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "{image1.png}",
  addOle: () => "{ole:oleObject1.bin}",
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

const OLE_BYTES = new Uint8Array([0xd0, 0xcf, 0x11, 0xe0]);
const PNG_BYTES = new Uint8Array([0x89, 0x50, 0x4e, 0x47]);

// Descriptor-level parse sees the {ole:…} placeholder as the raw r:id — map it
// the way the compiler would map the rewritten relationship id.
const readCtxWithEmbedding = {
  resolveRelationship: (rId: string) => {
    if (rId === "{ole:oleObject1.bin}") return "../embeddings/oleObject1.bin";
    if (rId === "{image1.png}") return "../media/image1.png";
    return undefined;
  },
  getPart: () => undefined,
  getRaw: (path: string) => {
    if (path === "../embeddings/oleObject1.bin") return OLE_BYTES;
    if (path === "../media/image1.png") return PNG_BYTES;
    return undefined;
  },
} as unknown as ReadContext;

function roundTrip(opts: OleOptions, ctx: ReadContext = readCtx) {
  const xml = oleDesc.stringify(opts, writeCtx)!;
  const doc = parseXml(xml);
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return { parsed: oleDesc.parse(el, ctx), xml };
}

describe("oleDesc round-trip", () => {
  it("registers embedded OLE binary and reads it back on parse", () => {
    const opts: OleOptions = {
      id: 100,
      name: "Test OLE",
      x: 50,
      y: 60,
      width: 200,
      height: 150,
      progId: "Excel.Sheet.12",
      embed: { data: OLE_BYTES, followColorScheme: "full" },
      iconImage: { data: PNG_BYTES, type: "png" },
    };
    const { parsed, xml } = roundTrip(opts, readCtxWithEmbedding);

    expect(xml).toContain('r:id="{ole:oleObject1.bin}"');
    // followColorScheme lives on p:embed, not on p:oleObj
    expect(xml).toContain('<p:embed followColorScheme="full"/>');
    expect(xml).toContain('<a:blip r:embed="{image1.png}"/>');
    expect(parsed.id).toBe(100);
    expect(parsed.name).toBe("Test OLE");
    expect(parsed.x).toBe(50);
    expect(parsed.y).toBe(60);
    expect(parsed.width).toBe(200);
    expect(parsed.height).toBe(150);
    expect(parsed.progId).toBe("Excel.Sheet.12");
    expect(parsed.embed).toBeDefined();
    expect(parsed.embed!.data).toEqual(OLE_BYTES);
    expect(parsed.embed!.followColorScheme).toBe("full");
    expect(parsed.iconImage).toBeDefined();
    expect(parsed.iconImage!.data).toEqual(PNG_BYTES);
    expect(parsed.iconImage!.type).toBe("png");
  });

  it("round-trips linked OLE object", () => {
    const opts: OleOptions = {
      id: 200,
      name: "Linked OLE",
      x: 10,
      y: 20,
      width: 300,
      height: 200,
      progId: "Word.Document.12",
      link: { rId: "rId3", autoUpdate: true },
    };
    const { parsed } = roundTrip(opts);

    expect(parsed.id).toBe(200);
    expect(parsed.name).toBe("Linked OLE");
    expect(parsed.progId).toBe("Word.Document.12");
    expect(parsed.link).toBeDefined();
    expect(parsed.link!.rId).toBe("rId3");
    expect(parsed.link!.autoUpdate).toBe(true);
  });

  it("round-trips linked OLE without autoUpdate", () => {
    const opts: OleOptions = {
      id: 250,
      link: { rId: "rId4" },
    };
    const { parsed } = roundTrip(opts);

    expect(parsed.link).toBeDefined();
    expect(parsed.link!.rId).toBe("rId4");
    expect(parsed.link!.autoUpdate).toBeUndefined();
  });

  it("round-trips OLE with showAsIcon", () => {
    const opts: OleOptions = {
      id: 300,
      name: "Icon OLE",
      showAsIcon: true,
      imageWidth: 64,
      imageHeight: 64,
      embed: { data: OLE_BYTES },
    };
    const { parsed } = roundTrip(opts);

    expect(parsed.id).toBe(300);
    expect(parsed.showAsIcon).toBe(true);
    expect(parsed.imageWidth).toBe(64);
    expect(parsed.imageHeight).toBe(64);
  });

  it("round-trips OLE with shapeId", () => {
    const opts: OleOptions = {
      id: 400,
      shapeId: "_x0000_s1025",
      embed: { data: OLE_BYTES },
    };
    const { parsed } = roundTrip(opts);

    expect(parsed.shapeId).toBe("_x0000_s1025");
  });

  it("round-trips OLE with followColorScheme", () => {
    const opts: OleOptions = {
      id: 500,
      embed: { data: OLE_BYTES, followColorScheme: "full" },
    };
    // followColorScheme round-trips inside embed, so the embedding bytes must
    // resolve for the embed object to be read back at all.
    const { parsed } = roundTrip(opts, readCtxWithEmbedding);

    expect(parsed.embed!.followColorScheme).toBe("full");
  });

  it("round-trips position with defaults", () => {
    const opts: OleOptions = {
      id: 700,
      embed: { data: OLE_BYTES },
    };
    const { parsed } = roundTrip(opts);

    expect(parsed.id).toBe(700);
    expect(parsed.name).toBe("Object 700");
    expect(parsed.x).toBe(0);
    expect(parsed.y).toBe(0);
    // default 100px = 952500 EMU
    expect(parsed.width).toBe(952500);
    expect(parsed.height).toBe(952500);
  });

  it("round-trips EMU conversion correctly", () => {
    const opts: OleOptions = {
      id: 800,
      x: 1024,
      y: 768,
      width: 1920,
      height: 1080,
      embed: { data: OLE_BYTES },
    };
    const { parsed } = roundTrip(opts);

    expect(parsed.x).toBe(1024);
    expect(parsed.y).toBe(768);
    expect(parsed.width).toBe(1920);
    expect(parsed.height).toBe(1080);
  });
});
