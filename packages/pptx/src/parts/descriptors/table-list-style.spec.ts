import type { ReadContext, WriteContext } from "@office-open/core/descriptor";
import { parse as parseXml } from "@office-open/xml";
import type { TableOptions } from "@shared/table/table-frame";
import { describe, expect, it } from "vite-plus/test";

import { tableDesc } from "./table";

const readContext: ReadContext = {
  resolveRelationship: () => undefined,
  getPart: () => undefined,
  getRaw: () => undefined,
};

class WriteContextStub {
  registerShapeId() {}
  addRelationship() {
    return "rId1";
  }
  addMedia() {
    return "";
  }
  addHyperlink() {}
  addImage() {}
}

describe("table cell list styles", () => {
  it("round-trips a list-style extension", () => {
    const options: TableOptions = {
      rows: [{ cells: [{ text: "Cell", listStyle: { ext: "" } }] }],
    };
    const xml = tableDesc.stringify(options, new WriteContextStub() as WriteContext)!;
    expect(xml).toContain("<a:lstStyle><a:extLst></a:extLst></a:lstStyle>");

    const element = parseXml(xml).elements?.[0];
    if (!element) throw new Error("fixture has no root element");
    const parsed = tableDesc.parse(element, readContext);
    const cell = parsed.rows?.[0]?.cells?.[0];
    expect(cell).toMatchObject({ text: "Cell", listStyle: { ext: "" } });
  });
});
