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
  const doc = parseXml(xml);
  const el = doc.elements![0];
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
    expect(result.externalBook?.definedNames![0].name).toBe("MyRange");
    expect(result.externalBook?.definedNames![0].refersTo).toBe("Sheet1!$A$1:$B$10");
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

    const rows = result.externalBook?.sheetDataSet![0].rows!;
    expect(rows).toHaveLength(1);
    expect(rows[0].rowNumber).toBe(1);
  });

  it("round-trips OLE link", () => {
    const opts: ExternalLinkOptions = {
      oleLink: {
        oleItems: [
          { name: "Item1", advise: true },
          { name: "Item2", prefer: true },
        ],
      },
      oleRId: "rId2",
    };
    const result = roundTrip(opts);

    expect(result.oleLink?.oleItems).toHaveLength(2);
    expect(result.oleLink?.oleItems![0].name).toBe("Item1");
  });
});
