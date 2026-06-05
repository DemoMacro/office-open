// Dialogsheet, Revision Log, QueryTable, and Metadata modules.
// These are standalone XML generators for specialized Excel features.

import { writeFileSync } from "node:fs";

import type { Context } from "@office-open/core";
import {
  Workbook,
  Packer,
  Dialogsheet,
  QueryTableXml,
  MetadataXml,
  RevisionHeadersXml,
  RevisionLogXml,
} from "@office-open/xlsx";

const wb = new Workbook({
  worksheets: [
    {
      name: "Data",
      rows: [
        { cells: [{ value: "Product" }, { value: "Price" }] },
        { cells: [{ value: "Widget" }, { value: 9.99 }] },
        { cells: [{ value: "Gadget" }, { value: 19.99 }] },
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(wb);
writeFileSync("My Workbook.xlsx", buffer);

// ── Standalone XML demos (not part of the workbook yet) ──

const context: Context = { fileData: wb, stack: [] };

// Dialogsheet
const dialogsheet = new Dialogsheet({
  name: "Dialog1",
  tabColor: "FF9900",
  pageMargins: { left: 0.5, right: 0.5, top: 0.75, bottom: 0.75 },
  pageSetup: { paperSize: 9, orientation: "portrait" },
  sheetProtection: { content: true },
});
console.log("Dialogsheet XML:", dialogsheet.toXml(context).slice(0, 200) + "...");

// QueryTable
const queryTable = new QueryTableXml({
  name: "QueryTable1",
  connectionId: 1,
  autoFormat: true,
  preserveFormatting: true,
  adjustColumnWidth: true,
  refreshOnLoad: true,
});
console.log("QueryTable XML:", queryTable.toXml(context).slice(0, 200) + "...");

// Metadata
const metadata = new MetadataXml({
  types: [
    {
      name: "cellMetadata",
      ghostRow: false,
      ghostCol: false,
      edit: true,
      delete: false,
      copy: true,
      paste: true,
    },
  ],
  strings: [{ value: "Active" }],
  futureMetadata: [{ name: "fxd", type: "cellMetadata" }],
});
console.log("Metadata XML:", metadata.toXml(context).slice(0, 200) + "...");

// Revision Log
const revisionHeaders = new RevisionHeadersXml({
  guid: "00000000-0000-0000-0000-000000000001",
  diskRevisions: true,
  headers: [
    {
      guid: "10000000-0000-0000-0000-000000000001",
      dateTime: "2026-06-04T12:00:00Z",
      userName: "TestUser",
      rId: "rId1",
      maxSheetId: 2,
      sheetIds: [{ id: 1 }, { id: 2 }],
    },
  ],
});
console.log("RevisionHeaders XML:", revisionHeaders.toXml(context).slice(0, 200) + "...");

const revisionLog = new RevisionLogXml([
  {
    type: "cellChange",
    data: { rId: 1, ref: "A1", sheetIndex: 0, oldValue: "Hello", newValue: "World" },
  },
  { type: "rowColumn", data: { action: "insertRow", rId: 2, row: 5, sheetIndex: 0 } },
  { type: "formatting", data: { rId: 3, ref: "B2:C3", sheetIndex: 0, s: 1 } },
]);
console.log("RevisionLog XML:", revisionLog.toXml(context).slice(0, 200) + "...");
