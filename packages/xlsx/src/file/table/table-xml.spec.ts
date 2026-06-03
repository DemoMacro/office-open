import { describe, expect, it } from "vite-plus/test";

import { TableXml, TotalsRowFunction } from "./table-xml";

const context = { stack: [] };

describe("TableXml", () => {
  it("generates minimal table XML with required attributes", () => {
    const table = new TableXml({
      id: 1,
      displayName: "Table1",
      ref: "A1:D10",
      columns: [{ name: "Name" }, { name: "Value" }, { name: "Category" }, { name: "Date" }],
    });
    const xml = table.toXml(context);

    expect(xml).toContain("<table ");
    expect(xml).toContain('id="1"');
    expect(xml).toContain('displayName="Table1"');
    expect(xml).toContain('ref="A1:D10"');
    expect(xml).toContain('name="Table1"');
    expect(xml).toContain("<tableColumns");
    expect(xml).toContain('count="4"');
    expect(xml).toContain('<tableColumn id="1" name="Name"/>');
    expect(xml).toContain("</table>");
  });

  it("generates table with style info", () => {
    const table = new TableXml({
      id: 2,
      displayName: "Sales",
      ref: "A1:C5",
      columns: [{ name: "Product" }, { name: "Amount" }, { name: "Region" }],
      style: {
        name: "TableStyleMedium2",
        showRowStripes: true,
        showColumnStripes: false,
      },
    });
    const xml = table.toXml(context);

    expect(xml).toContain('name="TableStyleMedium2"');
    expect(xml).toContain('showRowStripes="1"');
  });

  it("uses default style when no style specified", () => {
    const table = new TableXml({
      id: 1,
      displayName: "T1",
      ref: "A1:B2",
      columns: [{ name: "A" }, { name: "B" }],
    });
    const xml = table.toXml(context);

    expect(xml).toContain("TableStyleMedium9");
  });

  it("generates table with totals row", () => {
    const table = new TableXml({
      id: 3,
      displayName: "Budget",
      ref: "A1:C11",
      columns: [
        { name: "Item" },
        { name: "Cost", totalsRowFunction: TotalsRowFunction.SUM },
        { name: "Note", totalsRowLabel: "Total" },
      ],
      totalsRowCount: 1,
    });
    const xml = table.toXml(context);

    expect(xml).toContain('totalsRowCount="1"');
    expect(xml).toContain('totalsRowFunction="sum"');
    expect(xml).toContain('totalsRowLabel="Total"');
  });

  it("generates table with auto-filter", () => {
    const table = new TableXml({
      id: 1,
      displayName: "T1",
      ref: "A1:B10",
      columns: [{ name: "X" }, { name: "Y" }],
      autoFilter: "A1:B10",
    });
    const xml = table.toXml(context);

    expect(xml).toContain('<autoFilter ref="A1:B10"/>');
  });

  it("generates table with calculated column formula", () => {
    const table = new TableXml({
      id: 1,
      displayName: "Calc",
      ref: "A1:C5",
      columns: [
        { name: "A" },
        { name: "B" },
        { name: "Result", calculatedColumnFormula: "[@A]*[@B]" },
      ],
    });
    const xml = table.toXml(context);

    expect(xml).toContain("<calculatedColumnFormula>[@A]*[@B]</calculatedColumnFormula>");
  });

  it("omits default attribute values", () => {
    const table = new TableXml({
      id: 1,
      displayName: "T1",
      ref: "A1:B2",
      columns: [{ name: "A" }, { name: "B" }],
    });
    const xml = table.toXml(context);

    // tableType defaults to "worksheet" — should not appear
    expect(xml).not.toContain("tableType");
    // headerRowCount defaults to 1 — should not appear
    expect(xml).not.toContain("headerRowCount");
    // totalsRowCount defaults to 0 — should not appear
    expect(xml).not.toContain("totalsRowCount");
  });
});
