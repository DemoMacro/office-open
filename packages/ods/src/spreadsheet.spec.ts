import { generateOcf, ODF_NAMESPACES } from "@office-open/odf";
import { describe, expect, it } from "vite-plus/test";

import { generateWorkbook, parseWorkbook } from "./index";

const CALCEXT_NAMESPACE =
  'xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"';
const LOEXT_NAMESPACE =
  'xmlns:loext="urn:org:documentfoundation:names:experimental:office:xmlns:loext:1.0"';

function packageWith(
  body: string,
  automaticStyles = "",
  commonStyles = "<office:styles/>",
): Uint8Array {
  const content = `<?xml version="1.0"?><office:document-content ${ODF_NAMESPACES} ${CALCEXT_NAMESPACE} ${LOEXT_NAMESPACE} office:version="1.3"><office:automatic-styles>${automaticStyles}</office:automatic-styles><office:body><office:spreadsheet>${body}</office:spreadsheet></office:body></office:document-content>`;
  const styles = `<?xml version="1.0"?><office:document-styles ${ODF_NAMESPACES} office:version="1.3">${commonStyles}</office:document-styles>`;
  return generateOcf("application/vnd.oasis.opendocument.spreadsheet", {
    "content.xml": content,
    "styles.xml": styles,
  });
}

async function roundTrip(source: Uint8Array) {
  const parsed = await parseWorkbook(source);
  return { parsed, regenerated: await parseWorkbook(await generateWorkbook(parsed)) };
}

describe("ODS spreadsheet semantics", () => {
  it("round-trips sheet-local named expressions and ISO null dates", async () => {
    const { parsed, regenerated } = await roundTrip(
      packageWith(
        `<table:calculation-settings><table:null-date table:value-type="date" table:date-value="1904-01-01"/></table:calculation-settings><table:table table:name="Revenue"><table:table-row><table:table-cell office:value-type="float" office:value="2"/></table:table-row><table:named-expressions><table:named-range table:name="Total" table:cell-range-address="Revenue.A1" table:base-cell-address="Revenue.A1" loext:hidden="true"/></table:named-expressions></table:table>`,
      ),
    );
    expect(parsed.calculation?.nullDate).toEqual({ year: 1904, month: 1, day: 1 });
    expect(parsed.definedNames).toEqual([
      { name: "Total", value: "Revenue.A1", localSheetId: 0, hidden: true },
    ]);
    expect(regenerated.calculation?.nullDate).toEqual({ year: 1904, month: 1, day: 1 });
    expect(regenerated.definedNames).toEqual(parsed.definedNames);
  });

  it("round-trips content validations through cell references", async () => {
    const { parsed, regenerated } = await roundTrip(
      packageWith(
        `<table:content-validations><table:content-validation table:name="valid" table:condition="of:cell-content-is-in-list(&quot;A&quot;;&quot;B&quot;)" table:allow-empty-cell="false" table:display-list="unsorted" table:base-cell-address="Sheet1.A1"><table:error-message table:display="true" table:message-type="stop" table:title="Invalid"><text:p>Choose A or B</text:p></table:error-message></table:content-validation></table:content-validations><table:table table:name="Sheet1"><table:table-row><table:table-cell table:content-validation-name="valid" office:value-type="string"><text:p>A</text:p></table:table-cell></table:table-row></table:table>`,
      ),
    );
    expect(parsed.worksheets?.[0]?.dataValidations).toMatchObject([
      {
        sqref: "A1",
        type: "custom",
        formula1: 'of:cell-content-is-in-list("A";"B")',
        allowBlank: false,
        showDropDown: false,
        showErrorMessage: true,
        errorStyle: "stop",
        errorTitle: "Invalid",
        error: "Choose A or B",
      },
    ]);
    expect(regenerated.worksheets?.[0]?.dataValidations).toEqual(
      parsed.worksheets?.[0]?.dataValidations,
    );
  });

  it("round-trips conditional expression formatting and its differential style", async () => {
    const commonStyles =
      '<office:styles><style:style style:name="Red" style:family="table-cell"><style:table-cell-properties fo:background-color="#ed1c24"/></style:style></office:styles>';
    const { parsed, regenerated } = await roundTrip(
      packageWith(
        `<table:table table:name="Sheet1"><table:table-row><table:table-cell office:value-type="float" office:value="42"/></table:table-row><calcext:conditional-formats><calcext:conditional-format calcext:target-range-address="Sheet1.A1"><calcext:condition calcext:apply-style-name="Red" calcext:value="formula-is([.A1]=42)" calcext:base-cell-address="Sheet1.A1"/></calcext:conditional-format></calcext:conditional-formats></table:table>`,
        "",
        commonStyles,
      ),
    );
    expect(parsed.worksheets?.[0]?.conditionalFormats).toMatchObject([
      {
        sqref: "A1",
        rules: [{ type: "expression", formulas: ["[.A1]=42"], priority: 1, dxfId: 0 }],
      },
    ]);
    expect(parsed.dxfs).toEqual([
      { fill: { type: "solid", foregroundColor: { rgb: "FFed1c24" } } },
    ]);
    expect(regenerated.worksheets?.[0]?.conditionalFormats?.[0]?.rules).toMatchObject([
      { type: "expression", formulas: ["[.A1]=42"], dxfId: 0 },
    ]);
  });

  it("round-trips database ranges and source-mode data pilot tables", async () => {
    const { parsed, regenerated } = await roundTrip(
      packageWith(
        `<table:table table:name="Data"><table:table-row><table:table-cell office:value-type="float" office:value="1"/></table:table-row></table:table><table:database-ranges><table:database-range table:name="Database" table:target-range-address="Data.A1:A10" table:contains-header="true"/></table:database-ranges><table:data-pilot-tables><table:data-pilot-table table:name="Pivot" table:target-range-address="Data.B1:B4"><table:source-cell-range table:cell-range-address="Data.A1:A10"/><table:data-pilot-field table:source-field-name="Region" table:orientation="row" table:function="auto"/><table:data-pilot-field table:source-field-name="Sales" table:orientation="data" table:function="sum"/></table:data-pilot-table></table:data-pilot-tables>`,
      ),
    );
    expect(parsed.worksheets?.[0]?.tables).toMatchObject([
      { displayName: "Database", name: "Database", ref: "A1:A10", columns: [] },
    ]);
    expect(parsed.worksheets?.[0]?.pivotTables).toMatchObject([
      {
        mode: "source",
        name: "Pivot",
        source: "A1:A10",
        sourceSheet: "Data",
        location: "B1",
        rows: ["Region"],
        data: [{ field: "Sales", summarize: "sum" }],
      },
    ]);
    expect(regenerated.worksheets?.[0]?.tables).toMatchObject([
      { displayName: "Database", ref: "A1:A10" },
    ]);
    expect(regenerated.worksheets?.[0]?.pivotTables).toMatchObject([
      { name: "Pivot", source: "A1:A10", rows: ["Region"] },
    ]);
  });

  it("round-trips sheet protection and consolidation", async () => {
    const { parsed, regenerated } = await roundTrip(
      packageWith(
        `<table:table table:name="Sheet1"><loext:table-protection loext:select-protected-cells="false" loext:insert-rows="false"/><table:table-row><table:table-cell/></table:table-row></table:table><table:consolidation table:function="sum" table:source-cell-range-addresses="Sheet1.A1:A2 Sheet1.B1:B2" table:target-cell-address="Sheet1.D1" table:use-labels="both" table:link-to-source-data="true"/>`,
      ),
    );
    expect(parsed.worksheets?.[0]?.protection).toMatchObject({
      sheet: true,
      selectLockedCells: true,
      insertRows: true,
    });
    expect(parsed.worksheets?.[0]?.dataConsolidate).toMatchObject({
      function: "sum",
      topLabels: true,
      leftLabels: true,
      link: true,
      refs: [
        { sheet: "Sheet1", ref: "A1:A2" },
        { sheet: "Sheet1", ref: "B1:B2" },
      ],
    });
    expect(regenerated.worksheets?.[0]?.protection).toMatchObject({ sheet: true });
    expect(regenerated.worksheets?.[0]?.dataConsolidate).toMatchObject({ function: "sum" });
  });

  it("round-trips a cell-anchored styled shape", async () => {
    const automaticStyles =
      '<style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#3366ff" draw:stroke="solid" svg:stroke-width="1pt" svg:stroke-color="#ff0000"/></style:style>';
    const { parsed, regenerated } = await roundTrip(
      packageWith(
        `<table:table table:name="Sheet1"><table:table-row><table:table-cell><draw:rect draw:name="Card" draw:style-name="gr1" draw:z-index="2" svg:x="0.25cm" svg:y="0.5cm" svg:width="2cm" svg:height="1cm"><text:p>Label</text:p></draw:rect></table:table-cell></table:table-row></table:table>`,
        automaticStyles,
      ),
    );
    expect(parsed.worksheets?.[0]?.shapes?.[0]).toMatchObject({
      name: "Card",
      col: 1,
      row: 1,
      zOrder: 2,
      extentCx: 720000,
      extentCy: 360000,
      properties: {
        geometry: "rect",
        fill: { type: "solid", color: "3366ff" },
        outline: { width: 12700, type: "solidFill", color: "ff0000" },
      },
      textBody: { text: "Label" },
    });
    expect(regenerated.worksheets?.[0]?.shapes?.[0]).toMatchObject({
      name: "Card",
      col: 1,
      row: 1,
      zOrder: 2,
      extentCx: 720000,
      extentCy: 360000,
      properties: { geometry: "rect" },
      textBody: { text: "Label" },
    });
  });
});
