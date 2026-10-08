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

const readCtx = (
  binaries: Record<string, { path: string; bytes: Uint8Array; relType?: "oleObject" | "package" }>,
) =>
  ({
    resolveRelationship: (rid: string) => binaries[rid]?.path,
    resolveEmbeddingType: (rid: string) => binaries[rid]?.relType,
    getPart: () => undefined,
    getRaw: (path: string) => {
      for (const b of Object.values(binaries)) if (b.path === path) return b.bytes;
      return undefined;
    },
  }) as unknown as ReadContext;

describe("objectDesc.parse", () => {
  it("captures the Word 2010 anchor extension id", () => {
    const doc = parseXml(
      `<w:object ${NS} xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ` +
        `w14:anchorId="291A48E0"><v:shape/></w:object>`,
    );
    const opts = objectDesc.parse(doc.elements![0]!, readCtx({}));
    expect(opts.w14AnchorId).toBe("291A48E0");
  });

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

  it("round-trips a universal-measure original size", () => {
    const doc = parseXml(
      `<w:object ${NS} w:dxaOrig="362.05pt" w:dyaOrig="146.40pt"><v:shape/></w:object>`,
    );
    const opts = objectDesc.parse(doc.elements![0]!, readCtx({}));
    expect(opts.dxaOrig).toBe("362.05pt");
    expect(opts.dyaOrig).toBe("146.40pt");
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

  it("captures the complete preview shape with source relationship ids", () => {
    const el = parseObjectXml(
      `<v:shape id="_x0000_i1025" type="#_x0000_t75" alt="preview" ` +
        `style="width:100pt;height:50pt">` +
        `<v:imagedata r:id="rId8" o:title=""/></v:shape>`,
    );
    const opts = objectDesc.parse(
      el,
      readCtx({ rId8: { path: "word/media/image1.emf", bytes: new Uint8Array([1]) } }),
    );
    expect(opts.previewShape).toMatchObject({ alt: "preview" });
    expect(opts.previewShape!.imagedata).toMatchObject({ relationshipId: "rId8" });
  });

  it("round-trips OLE VML fillcolor, full style, and textbox", () => {
    const el = parseObjectXml(
      `<v:shape id="_x0000_i1025" type="#_x0000_t75" fillcolor="#ff0000" ` +
        `style="width:414pt;height:123.1pt;position:absolute;mso-wrap-style:square">` +
        `<v:textbox inset="2pt,2pt,2pt,2pt" o:insetmode="custom"/>` +
        `</v:shape>`,
    );
    const opts = objectDesc.parse(el, readCtx({}));
    expect(opts).toMatchObject({
      shapeId: "_x0000_i1025",
      fillcolor: "#ff0000",
      style: {
        width: "414pt",
        height: "123.1pt",
        position: "absolute",
        wrapStyle: "square",
      },
      textbox: { inset: "2pt,2pt,2pt,2pt", insetmode: "custom" },
    });
    const xml = objectDesc.stringify(opts, {} as never)!;
    expect(xml).toContain('fillcolor="#ff0000"');
    expect(xml).toContain(
      'style="width:414pt;height:123.1pt;position:absolute;mso-wrap-style:square"',
    );
    expect(xml).toContain('<v:textbox inset="2pt,2pt,2pt,2pt" o:insetmode="custom"/>');
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

  it("emits the Word 2010 anchor extension id", () => {
    const xml = objectDesc.stringify({ w14AnchorId: "291A48E0" }, writeCtx)!;
    expect(xml).toContain('w14:anchorId="291A48E0"');
  });

  it("rebuilds the mc:AlternateContent wrapper with the carried fallback", () => {
    const xml = objectDesc.stringify(
      {
        shapeId: "_x0000_i1025",
        width: "100pt",
        height: "50pt",
        mcChoiceRequires: "v",
        mcFallback: "<mc:Fallback><w:object><w:drawing/></w:object></mc:Fallback>",
        mcFallbackMedia: [{ fileName: "image1.png", data: new Uint8Array([1]), type: "png" }],
      },
      writeCtx,
    )!;
    expect(xml).toContain('<mc:AlternateContent><mc:Choice Requires="v">');
    expect(xml).toContain(
      "</mc:Choice><mc:Fallback><w:object><w:drawing/></w:object></mc:Fallback></mc:AlternateContent>",
    );
  });

  it("remaps fallback media placeholders to registered file names", () => {
    const renamingCtx = {
      file: {
        media: {
          addMedia: (_data: Uint8Array, _type: string, _factory: unknown, requested?: string) => ({
            fileName: requested === "image1.png" ? "image2.png" : requested,
          }),
        },
        embeddings: {
          addEmbedding: (_data: Uint8Array, requestedName?: string) => ({
            fileName: requestedName ?? "oleObject1.bin",
          }),
        },
      },
    } as unknown as BodyContext;
    const xml = objectDesc.stringify(
      {
        shapeId: "_x0000_i1025",
        mcFallback: '<mc:Fallback><a:blip r:embed="{image1.png}"/></mc:Fallback>',
        mcFallbackMedia: [{ fileName: "image1.png", data: new Uint8Array([1]), type: "png" }],
      },
      renamingCtx,
    )!;
    expect(xml).toContain("{image2.png}");
    expect(xml).not.toContain("{image1.png}");
  });

  it("remaps preview media relationships inside source VML", () => {
    const xml = objectDesc.stringify(
      {
        iconImage: { data: new Uint8Array([1]), type: "emf", fileName: "image1.emf" },
        previewShape: {
          alt: "preview",
          imagedata: { relationshipId: "rId8" },
          sourceXml: '<v:shape alt="preview"><v:imagedata r:id="rId8" o:title=""/></v:shape>',
        },
      },
      writeCtx,
    )!;
    expect(xml).toContain('alt="preview"');
    expect(xml).toContain('r:id="{image1.png}"');
    expect(xml).not.toContain('r:id="rId8"');
  });
});

describe("objectDesc.parse embedding relationship type", () => {
  it("captures a package-typed embedding rel for native-format parts", () => {
    // Word relates an embedded workbook as an OPC package, not an OLE
    // compound; the distinction must survive for the rel to re-emit correctly.
    const el = parseObjectXml(
      `<o:OLEObject Type="Embed" ProgID="Excel.Sheet.12" ShapeID="_x0000_i1025" ` +
        `DrawAspect="Content" ObjectID="_1" r:id="rId5"/>`,
    );
    const opts = objectDesc.parse(
      el,
      readCtx({
        rId5: {
          path: "word/embeddings/Book1.xlsx",
          bytes: new Uint8Array([9]),
          relType: "package",
        },
      }),
    );
    expect(opts.embed!.relationshipType).toBe("package");
  });

  it("captures an oleObject-typed embedding rel", () => {
    const el = parseObjectXml(
      `<o:OLEObject Type="Embed" ProgID="Package" ShapeID="_x0000_i1025" ` +
        `DrawAspect="Content" ObjectID="_2" r:id="rId6"/>`,
    );
    const opts = objectDesc.parse(
      el,
      readCtx({
        rId6: {
          path: "word/embeddings/oleObject1.bin",
          bytes: new Uint8Array([8]),
          relType: "oleObject",
        },
      }),
    );
    expect(opts.embed!.relationshipType).toBe("oleObject");
  });
});
