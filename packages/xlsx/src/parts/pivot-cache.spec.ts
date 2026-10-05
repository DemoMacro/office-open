import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { expect, it } from "vite-plus/test";

import {
  pivotCacheDefDesc,
  pivotCacheRecordsDesc,
  type PivotCacheDefDescriptorOptions,
  type PivotCacheRecordsDescriptorOptions,
} from "./pivot-cache";
import { XlsxParseError } from "./pivot/pivot-xml";

const writeCtx = { addRelationship: () => "rId1", addMedia: () => "" } as unknown as WriteContext;
const readCtx = {
  currentPart: "xl/pivotCache/pivotCacheDefinition1.xml",
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

const sourceData = {
  fieldNames: ["Name", "Amount"],
  records: [
    ["Alice", 100],
    ["Bob", 200],
  ],
};

function parseRoot(xml: string) {
  const el = parseXml(xml).elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return el;
}

it("keeps the source authoring cache definition path", () => {
  const opts: PivotCacheDefDescriptorOptions = {
    sourceRef: "A1:B3",
    sourceSheet: "Sheet1",
    sourceData,
    recordsRid: "rId1",
  };
  const xml = pivotCacheDefDesc.stringify(opts, writeCtx)!;
  const result = pivotCacheDefDesc.parse(parseRoot(xml), readCtx);
  expect(result.cacheSource?.children?.[0]?.attributes?.ref).toBe("A1:B3");
  expect(result.cacheSource?.children?.[0]?.attributes?.sheet).toBe("Sheet1");
  expect(result.cacheFields?.fields.map((field) => field.name)).toEqual(["Name", "Amount"]);
});

it("round-trips paired definition and records with OLAP metadata", () => {
  const definitionXml =
    `<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" recordCount="1" refreshedBy="Excel User" createdVersion="6">` +
    '<cacheSource type="external" connectionId="7"/><cacheFields count="2">' +
    '<cacheField name="Customer" numFmtId="0"><sharedItems count="1"><s v="Northwind"/></sharedItems></cacheField>' +
    '<cacheField name="Sales"><sharedItems containsSemiMixedTypes="0" containsString="0" containsNumber="1" minValue="12" maxValue="34"><n v="12"/><n v="34"/></sharedItems></cacheField>' +
    '</cacheFields><cacheHierarchies count="1"><cacheHierarchy uniqueName="[Customer].[Geo]" caption="Geo" count="0"/></cacheHierarchies>' +
    '<kpis count="1"><kpi uniqueName="[Measures].[Margin]" value="[Measures].[Sales]" goal="[Measures].[Target]"/></kpis>' +
    '<tupleCache><entries count="1"><e n="0" d="0"/></entries><sets count="1"><set countOfItems="1" rank="1" displayName="Top Customers"><tpls count="1"><tpl v="0"/></tpls></set></sets>' +
    '<serverFormats count="1"><serverFormat culture="en-US" format="Currency"/></serverFormats><queryCache count="1"><query t="Top Customers" c="1" r="1"><tpls count="1"><tpl f="1"/></tpls></query></queryCache></tupleCache>' +
    "</pivotCacheDefinition>";
  const recordsXml =
    '<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1"><r>' +
    '<x v="0" u="1" f="1"/><n v="12"/><b v="1"/><e v="#N/A"/><s v="Sales"/><d v="2024-01-01T00:00:00Z"/><m/>' +
    "</r></pivotCacheRecords>";

  const definition = pivotCacheDefDesc.parse(parseRoot(definitionXml), readCtx);
  const records = pivotCacheRecordsDesc.parse(parseRoot(recordsXml), readCtx);
  expect(definition.cacheFields?.fields[1]?.sharedItemsAttributes).toMatchObject({
    containsSemiMixedTypes: false,
    containsString: false,
    containsNumber: true,
    minValue: 12,
    maxValue: 34,
  });
  expect(records.records[0]?.entries.map((entry) => entry.type)).toEqual([
    "sharedItem",
    "number",
    "boolean",
    "error",
    "string",
    "date",
    "missing",
  ]);
  expect(pivotCacheDefDesc.stringify({ definition }, writeCtx)).toBe(definitionXml);
  expect(pivotCacheRecordsDesc.stringify({ records }, writeCtx)).toBe(recordsXml);
});

it("keeps the source authoring records path", () => {
  const opts: PivotCacheRecordsDescriptorOptions = { sourceData };
  const xml = pivotCacheRecordsDesc.stringify(opts, writeCtx)!;
  const result = pivotCacheRecordsDesc.parse(parseRoot(xml), readCtx);
  expect(result.records[0]?.entries).toEqual([
    { type: "sharedItem", value: 0, tuples: [], indexes: [] },
    { type: "number", value: 100, tuples: [], indexes: [] },
  ]);
});

it("reports unknown pivot cache XML children structurally", () => {
  const xml =
    '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheFields><notACacheField/></cacheFields></pivotCacheDefinition>';
  try {
    pivotCacheDefDesc.parse(parseRoot(xml), readCtx);
    throw new Error("expected parse failure");
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    const parseError = error as XlsxParseError;
    expect(parseError.part).toBe("xl/pivotCache/pivotCacheDefinition1.xml");
    expect(parseError.path).toBe("/pivotCacheDefinition/cacheFields");
    expect(parseError.name).toBe("notACacheField");
    expect(parseError.reason).toContain("unexpected child");
  }
});

it("round-trips calculated items and members", () => {
  const options: PivotCacheDefDescriptorOptions = {
    sourceRef: "A1:B3",
    sourceSheet: "Sheet1",
    sourceData: { fieldNames: ["Name"], records: [["A"]] },
    recordsRid: "rId1",
    cacheDefOpts: {
      calculatedItems: [
        {
          field: 0,
          formula: "A*2",
          pivotArea: { field: 0, references: [{ field: 0, x: [1] }] },
        },
      ],
      calculatedMembers: [
        {
          name: "Total",
          mdx: "[Measures].[Total]",
          memberName: "[Measures].[Total Member]",
          solveOrder: 2,
          set: true,
        },
      ],
    },
  };
  const xml = pivotCacheDefDesc.stringify(options, writeCtx)!;
  const parsed = pivotCacheDefDesc.parse(parseXml(xml).elements![0]!, readCtx);
  expect(parsed.calculatedItems).toEqual(options.cacheDefOpts?.calculatedItems);
  expect(parsed.calculatedMembers).toEqual(options.cacheDefOpts?.calculatedMembers);
  const reparsed = pivotCacheDefDesc.parse(
    parseXml(pivotCacheDefDesc.stringify({ definition: parsed }, writeCtx)!).elements![0]!,
    readCtx,
  );
  expect(reparsed.calculatedItems).toEqual(parsed.calculatedItems);
  expect(reparsed.calculatedMembers).toEqual(parsed.calculatedMembers);
});
