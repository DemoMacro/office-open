import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { metadataDesc } from "./metadata";
import type { MetadataOptions } from "./metadata";

const writeCtx = {} as unknown as WriteContext;
const readCtx = {} as unknown as ReadContext;

function roundTrip(opts: MetadataOptions) {
  const xml = metadataDesc.stringify(opts, writeCtx);
  if (!xml) throw new Error("stringify produced no XML");
  const doc = parseXml(xml, { nativeTypeAttributes: true });
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return metadataDesc.parse(el, readCtx);
}

describe("metadataDesc", () => {
  it("round-trips metadata types with the full copy/paste flag set", () => {
    const parsed = roundTrip({
      types: [
        {
          name: "XLDAPROPERTY",
          minSupportedVersion: 1,
          ghostRow: true,
          ghostCol: true,
          edit: true,
          delete: true,
          copy: true,
          pasteAll: true,
          pasteFormulas: true,
          pasteValues: true,
          pasteFormats: true,
          pasteComments: true,
          pasteDataValidation: true,
          pasteBorders: true,
          pasteColWidths: true,
          pasteNumberFormats: true,
          merge: true,
          splitFirst: true,
          splitAll: true,
          rowColShift: true,
          clearAll: true,
          clearFormats: true,
          clearContents: true,
          clearComments: true,
          assign: true,
          coerce: true,
          adjust: true,
          cellMeta: true,
        },
      ],
      strings: [{ value: "member name" }, { value: "caption" }],
    });
    const t = parsed.types![0]!;
    expect(t.name).toBe("XLDAPROPERTY");
    expect(t.minSupportedVersion).toBe(1);
    expect(t.ghostRow).toBe(true);
    expect(t.pasteDataValidation).toBe(true);
    expect(t.rowColShift).toBe(true);
    expect(t.coerce).toBe(true);
    expect(t.cellMeta).toBe(true);
    expect(parsed.strings).toEqual([{ value: "member name" }, { value: "caption" }]);
  });

  it("round-trips mdx tuple, set, member property, and kpi variants", () => {
    const parsed = roundTrip({
      mdx: [
        {
          functionType: "m",
          stringIndex: 0,
          tuple: {
            count: 2,
            culture: "en-US",
            styleIndex: 1,
            formatIndex: 2,
            backgroundColor: "FFAAAAAA",
            foregroundColor: "FF0000FF",
            italic: true,
            underline: true,
            strikethrough: true,
            bold: true,
            stringIndexes: [{ x: 0, show: true }, { x: 1 }],
          },
        },
        { functionType: "s", stringIndex: 1, set: { namespaceCount: 3, count: 4, order: "na" } },
        { functionType: "p", stringIndex: 2, memberProp: { nameIndex: 0, namePairIndex: 1 } },
        {
          functionType: "k",
          stringIndex: 3,
          kpi: { nameIndex: 0, namePairIndex: 1, property: "g" },
        },
      ],
    });
    const [tuple, set, memberProp, kpi] = parsed.mdx!;
    expect(tuple!.functionType).toBe("m");
    expect(tuple!.tuple).toMatchObject({
      count: 2,
      culture: "en-US",
      styleIndex: 1,
      formatIndex: 2,
      backgroundColor: "FFAAAAAA",
      foregroundColor: "FF0000FF",
      italic: true,
      underline: true,
      strikethrough: true,
      bold: true,
    });
    expect(tuple!.tuple!.stringIndexes).toEqual([{ x: 0, show: true }, { x: 1 }]);
    expect(set!.set).toMatchObject({ namespaceCount: 3, count: 4, order: "na" });
    expect(memberProp!.memberProp).toEqual({ nameIndex: 0, namePairIndex: 1 });
    expect(kpi!.kpi).toEqual({ nameIndex: 0, namePairIndex: 1, property: "g" });
  });

  it("round-trips future metadata blocks and cell/value metadata records", () => {
    const parsed = roundTrip({
      futureMetadata: [{ name: "XLDAPROPERTY", blocks: [{}, {}] }],
      cellMetadata: [
        {
          records: [
            { typeIndex: 0, valueIndex: 0 },
            { typeIndex: 1, valueIndex: 2 },
          ],
        },
      ],
      valueMetadata: [{ records: [{ typeIndex: 0, valueIndex: 1 }] }],
    });
    expect(parsed.futureMetadata![0]).toMatchObject({ name: "XLDAPROPERTY" });
    expect(parsed.futureMetadata![0]!.blocks).toHaveLength(2);
    expect(parsed.cellMetadata![0]!.records).toEqual([
      { typeIndex: 0, valueIndex: 0 },
      { typeIndex: 1, valueIndex: 2 },
    ]);
    expect(parsed.valueMetadata![0]!.records).toEqual([{ typeIndex: 0, valueIndex: 1 }]);
  });
});
