import type { ReadContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { BodyContext } from "../../context";
import { objectDesc } from "./object-element";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml" ' +
  'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function parseObjectXml(inner: string) {
  const doc = parseXml(`<w:object ${NS}>${inner}</w:object>`);
  return doc.elements![0]!;
}

const readCtx = (binaries: Record<string, { path: string; bytes: Uint8Array }>) =>
  ({
    resolveRelationship: (rid: string) => binaries[rid]?.path,
    getPart: () => undefined,
    getRaw: (path: string) => {
      for (const b of Object.values(binaries)) if (b.path === path) return b.bytes;
      return undefined;
    },
  }) as unknown as ReadContext;

describe("objectDesc.parse", () => {
  it("captures the v:shapetype preamble structurally", () => {
    const el = parseObjectXml(
      `<v:shapetype id="_x0000_t75" coordsize="21600,21600" o:spt="75" path="m@4@5l@4@11@9@11@9@5xe">` +
        `<v:formulas><v:f eqn="if lineDrawn pixelLineWidth 0"/></v:formulas>` +
        `</v:shapetype>`,
    );
    const opts = objectDesc.parse(el, readCtx({}));
    expect(opts.shapetype).toBeDefined();
    expect(opts.shapetype!.id).toBe("_x0000_t75");
    expect(opts.shapetype!.spt).toBe(75);
    expect(opts.shapetype!.formulas!.equations).toEqual(["if lineDrawn pixelLineWidth 0"]);
  });

  it("fetches icon and OLE binaries through the part rels", () => {
    const iconBytes = new Uint8Array([1, 2, 3]);
    const oleBytes = new Uint8Array([4, 5, 6, 7]);
    const el = parseObjectXml(
      `<v:shape id="_x0000_i1025" type="#_x0000_t75" style="width:414pt;height:123.1pt">` +
        `<v:imagedata r:id="rId8" o:title=""/>` +
        `</v:shape>` +
        `<o:OLEObject Type="Embed" ProgID="Visio.Drawing.11" ShapeID="_x0000_i1025" ` +
        `DrawAspect="Content" ObjectID="_1239361469" r:id="rId9"/>`,
    );
    const opts = objectDesc.parse(
      el,
      readCtx({
        rId8: { path: "word/media/image1.emf", bytes: iconBytes },
        rId9: { path: "word/embeddings/oleObject1.bin", bytes: oleBytes },
      }),
    );
    expect(opts.shapeId).toBe("_x0000_i1025");
    expect(opts.width).toBe("414pt");
    expect(opts.iconImage).toMatchObject({ data: iconBytes, type: "emf" });
    expect(opts.embed).toMatchObject({
      progId: "Visio.Drawing.11",
      objectId: "_1239361469",
      data: oleBytes,
    });
  });
});

describe("objectDesc.stringify", () => {
  const writeCtx = {
    file: {
      media: { addMedia: () => ({ fileName: "image1.png" }) },
      embeddings: {
        addEmbedding: (_data: Uint8Array, requestedName?: string) => ({
          fileName: requestedName ?? "oleObject1.bin",
        }),
      },
    },
  } as unknown as BodyContext;

  it("emits v:shapetype before the preview v:shape", () => {
    const xml = objectDesc.stringify(
      {
        shapetype: { id: "_x0000_t75", coordsize: "21600,21600", spt: 75 },
        shapeId: "_x0000_i1025",
        width: "100pt",
        height: "50pt",
        embed: { data: new Uint8Array([1]), progId: "Excel.Sheet.12" },
      },
      writeCtx,
    )!;
    const stIdx = xml.indexOf("<v:shapetype");
    const shapeIdx = xml.indexOf("<v:shape ");
    expect(stIdx).toBeGreaterThanOrEqual(0);
    expect(stIdx).toBeLessThan(shapeIdx);
    expect(xml).toContain('o:spt="75"');
  });
});
