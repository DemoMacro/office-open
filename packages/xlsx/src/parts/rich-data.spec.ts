import type { CustomDescriptor, ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import {
  parseRichDataPartRelationships,
  richValueDataDesc,
  richValueRelsDesc,
  richValueStructuresDesc,
  richValueTypesInfoDesc,
  type RichValueDataOptions,
  type RichValueRelsOptions,
  type RichValueStructuresOptions,
  type RichValueTypesInfoOptions,
} from "./rich-data";

const writeCtx = {} as unknown as WriteContext;
const readCtx = {} as unknown as ReadContext;

const roundTrip = <T>(descriptor: CustomDescriptor<T, WriteContext>, options: T): T => {
  const xml = descriptor.stringify(options, writeCtx)!;
  const root = parseXml(xml, { nativeTypeAttributes: true }).elements?.[0];
  expect(root).toBeDefined();
  return descriptor.parse(root!, readCtx);
};

const options: RichValueDataOptions = {
  count: 2,
  values: [
    { structureIndex: 0, fallback: { type: "s", text: "cached" }, values: ["1", "2"] },
    { structureIndex: 1, values: [] },
  ],
};

describe("richValueDataDesc", () => {
  it("preserves an empty value list", () => {
    expect(roundTrip(richValueDataDesc, { values: [] })).toEqual({ values: [] });
  });

  it("round-trips values, fallbacks, and extensions", () => {
    const withExtensions: RichValueDataOptions = {
      ...options,
      extensions: [{ uri: "https://example.com/ext", innerXml: "<x:test/>" }],
    };
    expect(roundTrip(richValueDataDesc, withExtensions)).toEqual(withExtensions);
  });

  it("rejects unknown children and attributes", () => {
    const unknownChild = parseXml(
      '<xlrd:rvData xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"><xlrd:nope/></xlrd:rvData>',
      { nativeTypeAttributes: true },
    ).elements?.[0];
    expect(() => richValueDataDesc.parse(unknownChild!, readCtx)).toThrowError(/unexpected child/);
    const unknownAttr = parseXml(
      '<xlrd:rvData xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" extra="1"/>',
      { nativeTypeAttributes: true },
    ).elements?.[0];
    expect(() => richValueDataDesc.parse(unknownAttr!, readCtx)).toThrowError(
      /unexpected attribute/,
    );
  });
});

describe("richValueStructuresDesc", () => {
  const structures: RichValueStructuresOptions = {
    structures: [
      {
        type: "_",
        keys: [
          { name: "_Self", type: "i" },
          { name: "_DisplayString", type: "s" },
        ],
      },
    ],
  };

  it("round-trips structures and keys", () => {
    expect(roundTrip(richValueStructuresDesc, structures)).toEqual(structures);
  });

  it("rejects unknown keys", () => {
    const root = parseXml(
      richValueStructuresDesc
        .stringify(structures, writeCtx)!
        .replace('<xlrd:k n="_Self" t="i"/>', "<xlrd:nope/>"),
      { nativeTypeAttributes: true },
    ).elements?.[0];
    expect(() => richValueStructuresDesc.parse(root!, readCtx)).toThrowError(/unexpected child/);
  });
});

describe("richValueTypesInfoDesc", () => {
  const typesInfo: RichValueTypesInfoOptions = {
    global: {
      keyFlags: { keys: [{ name: "_Self", flags: [{ name: "_Ignore", value: false }] }] },
    },
    types: {
      types: [
        {
          name: "Image",
          keyFlags: { keys: [{ name: "_Self", flags: [{ name: "_Type", value: true }] }] },
        },
      ],
    },
  };

  it("round-trips global and named types", () => {
    expect(roundTrip(richValueTypesInfoDesc, typesInfo)).toEqual(typesInfo);
  });
});

describe("richValueRelsDesc", () => {
  const rels: RichValueRelsOptions = { relationships: [{ rId: "rId1" }, { rId: "rId2" }] };

  it("round-trips relationship references", () => {
    expect(roundTrip(richValueRelsDesc, rels)).toEqual(rels);
  });

  it("rejects a missing r:id", () => {
    const root = parseXml(
      '<xlrvrel:richValueRels xmlns:xlrvrel="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel"><xlrvrel:rel/></xlrvrel:richValueRels>',
      { nativeTypeAttributes: true },
    ).elements?.[0];
    expect(() => richValueRelsDesc.parse(root!, readCtx)).toThrowError(/missing required value/);
  });
});

describe("parseRichDataPartRelationships", () => {
  it("parses internal and external relationships", () => {
    const root = parseXml(
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
        '<Relationship Id="rId1" Type="https://example.com/image" Target="../media/image1.png"/>' +
        '<Relationship Id="rId2" Type="https://example.com/hyperlink" Target="https://example.com/" TargetMode="External"/>' +
        "</Relationships>",
      { nativeTypeAttributes: true },
    ).elements?.[0];
    expect(parseRichDataPartRelationships(root)).toEqual([
      { rId: "rId1", relationshipType: "https://example.com/image", target: "../media/image1.png" },
      {
        rId: "rId2",
        relationshipType: "https://example.com/hyperlink",
        target: "https://example.com/",
        targetMode: "External",
      },
    ]);
  });
});
