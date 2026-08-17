import type { ReadContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import type { BodyContext } from "../../context";
import { parsePict, stringifyPict } from "./pict";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml" ' +
  'xmlns:o="urn:schemas-microsoft-com:office:office" ' +
  'xmlns:w10="urn:schemas-microsoft-com:office:word" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';

function parsePictXml(inner: string) {
  const doc = parseXml(`<w:pict ${NS}>${inner}</w:pict>`);
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

describe("parsePict", () => {
  it("keeps shapetype + shape children ordered", () => {
    const el = parsePictXml(
      `<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800"/>` +
        `<v:shape id="_x0000_s1026" type="#_x0000_t136" style="width:120pt;height:24pt" fillcolor="#369">` +
        `<v:shadow on="t" type="perspective"/>` +
        `</v:shape>`,
    );
    const opts = parsePict(el, readCtx({}));
    expect(opts.children).toHaveLength(2);
    expect(opts.children![0]).toMatchObject({ shapetype: { id: "_x0000_t136", spt: 136 } });
    expect(opts.children![1]).toMatchObject({
      shape: {
        id: "_x0000_s1026",
        type: "#_x0000_t136",
        shadow: { on: true, type: "perspective" },
      },
    });
  });

  it("bridges v:imagedata r:id media to a {fileName} placeholder", () => {
    const bytes = new Uint8Array([9, 9, 9]);
    const el = parsePictXml(
      `<v:shape id="_x0000_i1027" type="#_x0000_t75" style="width:42.75pt;height:57.75pt">` +
        `<v:imagedata r:id="rId8" o:title=""/>` +
        `</v:shape>`,
    );
    const opts = parsePict(el, readCtx({ rId8: { path: "word/media/image1.wmf", bytes } }));
    expect(opts.children![0]).toMatchObject({
      shape: { imagedata: { relationshipId: "{image1.wmf}", officeTitle: "" } },
    });
    expect(opts.media).toEqual([{ fileName: "image1.wmf", data: bytes, type: "wmf" }]);
  });

  it("bridges imagedata nested inside a v:group", () => {
    const bytes = new Uint8Array([1]);
    const el = parsePictXml(
      `<v:group style="width:100pt;height:50pt">` +
        `<v:shape id="_x0000_s1027" style="width:40pt;height:20pt">` +
        `<v:imagedata r:id="rId2"/>` +
        `</v:shape>` +
        `</v:group>`,
    );
    const opts = parsePict(el, readCtx({ rId2: { path: "word/media/image9.png", bytes } }));
    expect(opts.children![0]).toMatchObject({
      group: { children: [{ shape: { imagedata: { relationshipId: "{image9.png}" } } }] },
    });
    expect(opts.media).toEqual([{ fileName: "image9.png", data: bytes, type: "png" }]);
  });

  it("leaves a dangling r:id verbatim instead of registering empty media", () => {
    const el = parsePictXml(`<v:shape id="_x0000_i1028"><v:imagedata r:id="rId404"/></v:shape>`);
    const opts = parsePict(el, readCtx({}));
    expect(opts.children![0]).toMatchObject({
      shape: { imagedata: { relationshipId: "rId404" } },
    });
    expect(opts.media).toBeUndefined();
  });
});

describe("stringifyPict", () => {
  const writeCtx = (renames: Record<string, string> = {}) => {
    const registrations: { data: Uint8Array; type: string; fileName?: string }[] = [];
    const ctx = {
      file: {
        media: {
          addMedia: (
            data: Uint8Array,
            type: string,
            build: (fileName: string) => unknown,
            fileName?: string,
          ) => {
            registrations.push({ data, type, fileName });
            const finalName =
              fileName !== undefined ? (renames[fileName] ?? fileName) : "image1.png";
            return { fileName: build(finalName) ? finalName : finalName };
          },
        },
      },
    } as unknown as BodyContext;
    return { ctx, registrations };
  };

  it("emits an empty w:pict for no children", () => {
    expect(stringifyPict({}, writeCtx().ctx)).toBe("<w:pict/>");
  });

  it("round-trips a WordArt shape with textpath", () => {
    const xml =
      `<v:shape id="_x0000_i1030" type="#_x0000_t136" style="width:227.55pt;height:22.4pt">` +
      `<v:shadow color="#868686"/>` +
      `<v:textpath style="font-family:&quot;Arial Black&quot;;font-size:16pt" trim="t" string="Title"/>` +
      `</v:shape>`;
    const opts = parsePict(parsePictXml(xml), readCtx({}));
    const out = stringifyPict(opts, writeCtx().ctx);
    expect(out).toContain('type="#_x0000_t136"');
    expect(out).toContain('<v:shadow color="#868686"/>');
    expect(out).toContain('string="Title"');
    expect(out.startsWith("<w:pict>")).toBe(true);
  });

  it("registers media and remaps the placeholder when dedup renames it", () => {
    const bytes = new Uint8Array([7, 7]);
    const { ctx, registrations } = writeCtx({ "image1.wmf": "image3.wmf" });
    const out = stringifyPict(
      {
        children: [
          {
            shape: { id: "_x0000_i1031", imagedata: { relationshipId: "{image1.wmf}" } },
          },
        ],
        media: [{ fileName: "image1.wmf", data: bytes, type: "wmf" }],
      },
      ctx,
    );
    expect(registrations).toEqual([{ data: bytes, type: "wmf", fileName: "image1.wmf" }]);
    expect(out).toContain('r:id="{image3.wmf}"');
  });

  it("skips empty media entries (OPC guard)", () => {
    const { ctx, registrations } = writeCtx();
    const out = stringifyPict(
      {
        children: [{ shape: { id: "_x0000_i1032", imagedata: { relationshipId: "{x.}" } } }],
        media: [{ fileName: "x.", data: new Uint8Array(0), type: "" }],
      },
      ctx,
    );
    expect(registrations).toHaveLength(0);
    expect(out).toContain('r:id="{x.}"');
  });
});
