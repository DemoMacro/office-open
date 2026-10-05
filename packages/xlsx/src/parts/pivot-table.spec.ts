import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import { expect, it } from "vite-plus/test";

import { pivotTableDesc, type PivotTableDescriptorOptions } from "./pivot-table";
import type { PivotSourceData } from "./pivot/pivot-utils";
import { XlsxParseError } from "./pivot/pivot-xml";

const writeCtx = { addRelationship: () => "rId1", addMedia: () => "" } as unknown as WriteContext;
const readCtx = {
  currentPart: "xl/pivotTables/pivotTable1.xml",
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
} as unknown as ReadContext;

const sourceData: PivotSourceData = {
  fieldNames: ["Region", "Product", "Sales"],
  records: [
    ["East", "A", 100],
    ["East", "B", 200],
    ["West", "A", 150],
    ["West", "B", 250],
  ],
};

function roundTrip(opts: PivotTableDescriptorOptions) {
  const xml = pivotTableDesc.stringify(opts, writeCtx)!;
  const el = parseXml(xml).elements?.[0];
  if (!el) throw new Error("parsed document has no root element");
  return { xml, result: pivotTableDesc.parse(el, readCtx) };
}

it("round-trips a complete pivotTable definition", () => {
  const xml =
    '<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="SalesPivot" cacheId="7" dataCaption="Values" applyNumberFormats="0" applyWidthHeightFactors="1">' +
    '<location ref="A3:D9" firstHeaderRow="1" firstDataRow="2" firstDataCol="1" rowPageCount="1" colPageCount="1"/>' +
    '<pivotFields count="3"><pivotField axis="axisRow" showAll="0"/><pivotField axis="axisCol" showAll="0"/><pivotField dataField="1" showAll="0"/></pivotFields>' +
    '<rowFields count="1"><field x="0"/></rowFields><rowItems count="2"><i><x/></i><i t="grand"><x/></i></rowItems>' +
    '<colFields count="1"><field x="1"/></colFields><colItems count="2"><i><x/></i><i t="grand"><x/></i></colItems>' +
    '<dataFields count="1"><dataField name="Sum of Sales" fld="2" baseField="0" baseItem="0"/></dataFields>' +
    '<pivotTableStyleInfo name="PivotStyleDark1" showRowHeaders="1" showColHeaders="1"/></pivotTableDefinition>';
  const definition = pivotTableDesc.parse(parseXml(xml).elements![0]!, readCtx);
  expect(definition.name).toBe("SalesPivot");
  expect(definition.cacheId).toBe(7);
  expect(definition.location?.attributes?.ref).toBe("A3:D9");
  expect(definition.pivotFields?.children?.[0]?.attributes?.axis).toBe("axisRow");
  expect(definition.rowItems?.children?.map((item) => item.name)).toEqual(["i", "i"]);
  expect(definition.dataFields?.children?.[0]?.attributes?.name).toBe("Sum of Sales");
  expect(pivotTableDesc.stringify({ definition }, writeCtx)).toBe(xml);
});

it("keeps source-mode generation and parses its complete definition", () => {
  const { result } = roundTrip({
    options: {
      mode: "source",
      source: "Sheet1!A1:C5",
      rows: ["Region", "Product"],
      columns: undefined,
      data: [{ field: "Sales", summarize: "average", name: "Avg Sales" }],
      style: "PivotStyleDark1",
    },
    sourceData,
    cacheId: 5,
  });
  expect(result.name).toBe("PivotTable1");
  expect(result.cacheId).toBe(5);
  expect(result.location?.attributes?.ref).toBeDefined();
  expect(result.pivotFields?.children?.[0]?.attributes?.axis).toBe("axisRow");
  expect(result.pivotFields?.children?.[1]?.attributes?.axis).toBe("axisRow");
  expect(result.dataFields?.children?.[0]?.attributes?.subtotal).toBe("average");
  expect(result.pivotTableStyleInfo?.attributes?.name).toBe("PivotStyleDark1");
});

it("reports unknown pivotTable children structurally", () => {
  const xml =
    '<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><location/><notAllowed/></pivotTableDefinition>';
  try {
    pivotTableDesc.parse(parseXml(xml).elements![0]!, readCtx);
    throw new Error("expected parse failure");
  } catch (error) {
    expect(error).toBeInstanceOf(XlsxParseError);
    const parseError = error as XlsxParseError;
    expect(parseError.part).toBe("xl/pivotTables/pivotTable1.xml");
    expect(parseError.path).toBe("/pivotTableDefinition");
    expect(parseError.name).toBe("notAllowed");
    expect(parseError.reason).toContain("unexpected child");
  }
});
