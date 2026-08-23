// Packs every CT_Worksheet child element the compiler can author into a single
// sheet, so XSD validation exercises the full element sequence — mis-ordered
// siblings only surface when they co-occur.
import { mkdirSync, writeFileSync } from "node:fs";

import { generateWorkbook } from "@office-open/xlsx";

const png1x1 = new Uint8Array([
  0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
  0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
  0x89, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x62, 0x00, 0x01, 0x00, 0x00,
  0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
  0x42, 0x60, 0x82,
]);

const buffer = await generateWorkbook({
  calcPr: { calcId: 191029 },
  oleSize: "A1:D10",
  // Excel requires every sheet-level customSheetView guid to have a
  // same-guid workbook-level customWorkbookView — orphaned sheet views
  // make Excel refuse the file outright.
  customWorkbookViews: [
    {
      name: "Snapshot",
      guid: "{11111111-2222-3333-4444-555555555555}",
      windowWidth: 1936,
      windowHeight: 1048,
      activeSheetId: 1,
    },
  ],
  worksheets: [
    {
      name: "Everything",
      sheetPr: { published: false, enableFormatConditionsCalculation: false },
      tabColor: { rgb: "FF4472C4" },
      dimension: "A1:F12",
      sheetView: { showGridLines: false, zoomScale: 90 },
      selection: [{ activeCell: "B2", sqref: "B2" }],
      sheetFormatPr: { defaultRowHeight: 16, outlineLevelRow: 1 },
      columns: [{ min: 1, max: 6, width: 12 }],
      rows: [
        // Header cells must cover every table column — Excel refuses the
        // file when a tableColumn name has no matching header cell value.
        {
          cells: [
            { value: "Header" },
            { value: "Column2" },
            { value: "Column3" },
            { value: "Column4" },
            { value: "Column5" },
            { value: "Column6" },
          ],
        },
        {
          rowNumber: 2,
          cells: [{ value: 1 }, { value: 2 }],
        },
      ],
      sheetCalcPr: { fullCalcOnLoad: true },
      protection: { sheet: true, formatCells: false },
      protectedRanges: [{ sqref: "A5:C7", name: "Locked" }],
      scenarios: { scenarios: [{ name: "Base", inputCells: [{ reference: "B2", val: 2 }] }] },
      // Sheet-level autoFilter must not overlap the table range below —
      // Excel refuses the file on any autoFilter×table range intersection.
      autoFilter: { ref: "H1:J12", columns: [{ colId: 0, filters: { values: ["Header"] } }] },
      dataConsolidate: { function: "sum", refs: ["A2:F12"] },
      customSheetViews: [{ guid: "{11111111-2222-3333-4444-555555555555}", scale: 80 }],
      mergeCells: [{ ref: "D5:E5" }],
      phoneticPr: { fontId: 1 },
      conditionalFormats: [
        { sqref: "F2:F12", rules: [{ type: "cellIs", operator: "greaterThan", formulas: ["10"] }] },
      ],
      dataValidations: [{ sqref: "B2:B12", type: "whole", operator: "greaterThan", formula1: "0" }],
      hyperlinks: [{ cell: "A12", url: "https://example.com" }],
      printOptions: { horizontalCentered: true, gridLines: true },
      pageMargins: { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 },
      pageSetup: { paperSize: 9, orientation: "portrait", usePrinterDefaults: false, copies: 1 },
      headerFooter: {
        oddHeader: "Header &A",
        oddFooter: "Page &P",
        scaleWithDoc: false,
        alignWithMargins: false,
      },
      rowBreaks: [{ id: 8, manual: true }],
      colBreaks: [{ id: 3, manual: true }],
      // customProperties (customPr name + r:id) requires a relationship the
      // worksheet's .rels must declare — round-trip only; XSD marks r:id required.
      cellWatches: [{ reference: "B2" }],
      ignoredErrors: [{ sqref: "A2:A12", numberStoredAsText: true }],
      oleObjects: [{ shapeId: 1, progId: "Pkg" }],
      webPublishItems: [
        {
          id: 1,
          divId: "div1",
          sourceType: "range",
          sourceRef: "A1:F12",
          destinationFile: "out.htm",
          title: "All",
        },
      ],
      images: [{ type: "png", data: png1x1, col: 1, row: 1, toCol: 3, toRow: 4 }],
      comments: [{ cell: "C3", author: "Auditor", text: "checked" }],
      tables: [
        {
          id: 1,
          name: "AllTable",
          displayName: "AllTable",
          ref: "A1:F2",
          columns: [
            { name: "Header" },
            { name: "Column2" },
            { name: "Column3" },
            { name: "Column4" },
            { name: "Column5" },
            { name: "Column6" },
          ],
        },
      ],
    },
  ],
});

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/33-worksheet-element-order.xlsx", buffer);
console.log("written .temp/33-worksheet-element-order.xlsx");
