import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import { externalLinkDesc } from "./external-link";
import type { ExternalLinkOptions } from "./external-link";

// ── Minimal context stubs ──

const writeCtx = {
  addRelationship: () => "rId1",
  addMedia: () => "",
} as unknown as WriteContext;

const readCtx = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

function roundTrip(opts: ExternalLinkOptions) {
  const xml = externalLinkDesc.stringify(opts, writeCtx)!;
  // nativeTypeAttributes mirrors the real xlsx parse path (ParsedArchive.get
  // coerces "1"/"0" to numbers), so reads are exercised against numeric
  // coercion rather than a permissive non-coerced parse.
  const doc = parseXml(xml, { nativeTypeAttributes: true });
  const el = doc.elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return externalLinkDesc.parse(el, readCtx) as unknown as ExternalLinkOptions;
}

// ── Tests ──

describe("externalLinkDesc round-trip", () => {
  it("round-trips external book with sheet names", () => {
    const opts: ExternalLinkOptions = {
      externalBook: {
        sheetNames: ["Sheet1", "Sheet2"],
      },
      bookRId: "rId1",
    };
    const result = roundTrip(opts);

    expect(result.bookRId).toBe("rId1");
    expect(result.externalBook?.sheetNames).toEqual(["Sheet1", "Sheet2"]);
  });

  it("round-trips external book with defined names", () => {
    const opts: ExternalLinkOptions = {
      externalBook: {
        sheetNames: ["Data"],
        definedNames: [
          { name: "MyRange", refersTo: "Sheet1!$A$1:$B$10" },
          { name: "Total", refersTo: "Sheet1!$C$1", sheetId: 1 },
        ],
      },
      bookRId: "rId1",
    };
    const result = roundTrip(opts);

    expect(result.externalBook?.definedNames).toHaveLength(2);
    expect(result.externalBook?.definedNames![0]?.name).toBe("MyRange");
    expect(result.externalBook?.definedNames![0]?.refersTo).toBe("Sheet1!$A$1:$B$10");
  });

  it("round-trips external book with sheet data", () => {
    const opts: ExternalLinkOptions = {
      externalBook: {
        sheetDataSet: [
          {
            sheetId: 1,
            rows: [
              {
                rowNumber: 1,
                cells: [
                  { reference: "A1", type: "s", value: "Hello" },
                  { reference: "B1", type: "n", value: "42" },
                ],
              },
            ],
          },
        ],
      },
      bookRId: "rId1",
    };
    const result = roundTrip(opts);

    const rows = result.externalBook!.sheetDataSet![0]!.rows!;
    expect(rows).toHaveLength(1);
    expect(rows[0]?.rowNumber).toBe(1);
  });

  it("round-trips OLE link", () => {
    const opts: ExternalLinkOptions = {
      oleLink: {
        progId: "Word.Document.12",
        oleItems: [
          { name: "Item1", advise: true },
          { name: "Item2", preferPic: true, icon: true },
        ],
      },
      oleRId: "rId2",
    };
    const result = roundTrip(opts);

    expect(result.oleLink?.progId).toBe("Word.Document.12");
    expect(result.oleLink?.oleItems).toHaveLength(2);
    expect(result.oleLink?.oleItems![0]?.name).toBe("Item1");
    expect(result.oleLink?.oleItems![1]?.preferPic).toBe(true);
    expect(result.oleLink?.oleItems![1]?.icon).toBe(true);
  });

  it("round-trips DDE link with cached values", () => {
    const opts: ExternalLinkOptions = {
      ddeLink: {
        ddeService: "Excel",
        ddeTopic: "Sheet1",
        ddeItems: [
          {
            name: "R1C1:R2C2",
            ole: true,
            advise: true,
            values: {
              rows: 2,
              cols: 2,
              values: [
                { value: "1", type: "n" },
                { value: "text", type: "str" },
              ],
            },
          },
          { name: "plain" },
        ],
      },
    };
    const result = roundTrip(opts);

    expect(result.ddeLink?.ddeService).toBe("Excel");
    expect(result.ddeLink?.ddeTopic).toBe("Sheet1");
    const [item1, item2] = result.ddeLink?.ddeItems ?? [];
    expect(item1?.name).toBe("R1C1:R2C2");
    expect(item1?.ole).toBe(true);
    expect(item1?.advise).toBe(true);
    expect(item1?.values?.rows).toBe(2);
    expect(item1?.values?.cols).toBe(2);
    expect(item1?.values?.values[0]).toEqual({ value: "1", type: "n" });
    expect(item1?.values?.values[1]).toEqual({ value: "text", type: "str" });
    expect(item2?.name).toBe("plain");
    expect(item2?.values).toBeUndefined();
  });

  it("round-trips boolean flags under nativeTypeAttributes coercion", () => {
    const opts: ExternalLinkOptions = {
      externalBook: {
        sheetNames: ["S1"],
        definedNames: [{ name: "N1", sheetId: 3 }],
        sheetDataSet: [{ sheetId: 1, refreshError: true }],
      },
      oleLink: { oleItems: [{ name: "I1", advise: true, preferPic: true }] },
    };
    const result = roundTrip(opts);

    const dn = result.externalBook?.definedNames![0]!;
    expect(dn.sheetId).toBe(3);
    expect(result.externalBook?.sheetDataSet![0]?.refreshError).toBe(true);
    expect(result.oleLink?.oleItems![0]?.advise).toBe(true);
    expect(result.oleLink?.oleItems![0]?.preferPic).toBe(true);
  });
});
