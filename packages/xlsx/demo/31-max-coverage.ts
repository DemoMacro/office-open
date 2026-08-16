// Combines every safe authoring domain in one workbook — formulas, rich text,
// validation, conditional formats, drawings, tables, pivot, scenarios — to
// surface combination defects the single-feature demos cannot hit.
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
  workbookPr: { defaultThemeVersion: 202300, date1904: false },
  calcPr: { calcId: 191029 },
  bookView: { xWindow: 120, yWindow: 120, windowWidth: 20000, windowHeight: 12000 },
  customProperties: [{ name: "Reviewed", value: true }],
  definedNames: [{ name: "_xlnm.Print_Area", value: "Data!$A$1:$D$8", localSheetId: 0 }],
  appProperties: { company: "Example", application: "office-open", appVersion: "1.0000" },
  dxfs: [{ font: { color: "9C0006", bold: true } }, { fill: { color: "C6EFCE" } }],
  cellStyles: [
    { name: "Comma", xfId: 0 },
    { name: "Accent", xfId: 0, builtinId: 40 },
  ],
  worksheets: [
    // ── Sheet 1: cells, formulas, rich text, merges, freezes, validation, cf ──
    {
      name: "Data",
      rows: [
        {
          height: 22,
          cells: [
            { value: "Region" },
            { value: "Product" },
            { value: "Revenue" },
            {
              value: {
                runs: [
                  { text: "Rich ", properties: { bold: true } },
                  { text: "header", properties: { italic: true, color: "008000" } },
                ],
              },
            },
          ],
        },
        { cells: [{ value: "North" }, { value: "Widget" }, { value: 1200 }, { value: 10 }] },
        { cells: [{ value: "South" }, { value: "Gadget" }, { value: 800 }, { value: 20 }] },
        { cells: [{ value: "East" }, { value: "Widget" }, { value: 1500 }, { value: 30 }] },
        { cells: [{ value: "West" }, { value: "Gadget" }, { value: 950 }, { value: 40 }] },
        {
          cells: [
            { value: "Total" },
            { formula: 'CONCATENATE(A2,"-",A5)' },
            { formula: "SUM(C2:C5)" },
            { formula: "SUM(D2:D5)" },
          ],
        },
        {
          cells: [
            { value: "Lookup" },
            { formula: 'VLOOKUP("South",A2:C5,3,FALSE)' },
            { formula: 'IF(C6>4000,"high","low")' },
          ],
        },
      ],
      mergeCells: [{ from: { row: 8, col: 1 }, to: { row: 8, col: 4 } }],
      freezePanes: { row: 1, col: 1 },
      autoFilter: { ref: "A1:D5" },
      hyperlinks: [
        {
          cell: "F2",
          target: { type: "external", url: "https://example.com" },
          tooltip: "Open example",
        },
        { cell: "F3", target: { type: "internal", location: "Drawing!A1" } },
      ],
      dataValidations: [
        { type: "list", formula1: '"Yes,No"', sqref: "G2:G10" },
        { type: "whole", operator: "between", formula1: "1", formula2: "100", sqref: "H2:H10" },
      ],
      conditionalFormats: [
        {
          sqref: "C2:C5",
          rules: [
            { type: "cellIs", operator: "greaterThan", formulas: ["1000"], priority: 1, dxfId: 0 },
            { type: "top10", rank: 1, priority: 2, dxfId: 0 },
          ],
        },
      ],
      columns: [
        { min: 1, max: 1, width: 12, customWidth: true },
        { min: 2, max: 2, width: 14, customWidth: true },
      ],
      comments: [
        {
          cell: "A1",
          author: "Max",
          text: "Header cell with a note.",
          visible: false,
          anchor: [0, 0, 0, 0, 2, 0, 3, 0],
        },
      ],
      pageMargins: { left: "0.7in", right: "0.7in", top: "0.9in", bottom: "0.9in" },
      headerFooter: { oddHeader: "&CMax sample" },
      protection: { password: "hint", sheet: true, selectLockedCells: false },
      sheetView: { showGridLines: false, zoomScale: 90 },
    },
    // ── Sheet 2: drawing objects — images, shapes, connectors, charts ──
    {
      name: "Drawing",
      rows: [{ cells: [{ value: "Anchored drawing objects" }] }],
      images: [
        { data: png1x1, type: "png", col: 1, row: 2 },
        { data: png1x1, type: "png", col: 3, row: 2 },
      ],
      shapes: [
        {
          col: 1,
          row: 6,
          toCol: 4,
          toRow: 10,
          name: "Rectangle",
          spPr: { geometry: "rect", fill: { type: "solid", color: "4472C4" } },
        },
        {
          col: 5,
          row: 6,
          toCol: 8,
          toRow: 10,
          name: "Ellipse",
          spPr: {
            geometry: "ellipse",
            fill: { type: "solid", color: "ED7D31" },
            outline: { width: 12700, color: "000000" },
          },
        },
      ],
      connectors: [
        {
          col: 4,
          row: 8,
          toCol: 5,
          toRow: 10,
          spPr: { geometry: "line" },
          locking: { noAdjustHandles: true },
        },
      ],
      charts: [
        {
          type: "column",
          title: "Revenue by region",
          categories: ["N", "S", "E", "W"],
          series: [{ name: "2026", values: [1200, 800, 1500, 950] }],
          col: 10,
          row: 2,
        },
      ],
    },
    // ── Sheet 3: sml table + pivot (safe authoring form, no cf) ──
    {
      name: "TableHost",
      rows: [
        { cells: [{ value: "City" }, { value: "Category" }, { value: "Revenue" }] },
        { cells: [{ value: "Beijing" }, { value: "Food" }, { value: 320 }] },
        { cells: [{ value: "Beijing" }, { value: "Tech" }, { value: 580 }] },
        { cells: [{ value: "Shanghai" }, { value: "Food" }, { value: 410 }] },
        { cells: [{ value: "Shanghai" }, { value: "Tech" }, { value: 720 }] },
        { cells: [{ value: "Shenzhen" }, { value: "Food" }, { value: 195 }] },
        { cells: [{ value: "Shenzhen" }, { value: "Tech" }, { value: 850 }] },
      ],
      tables: [
        {
          id: 1,
          displayName: "SalesTable",
          name: "SalesTable",
          ref: "A1:C7",
          autoFilter: "A1:C7",
          columns: [{ name: "City" }, { name: "Category" }, { name: "Revenue" }],
          style: { name: "TableStyleMedium9", showRowStripes: true },
        },
      ],
    },
    {
      name: "Pivot",
      pivotTables: [
        {
          name: "MaxPivot",
          source: "A1:C7",
          sourceSheet: "TableHost",
          location: "A3",
          rows: ["City"],
          data: [{ field: "Revenue", summarize: "sum" }],
        },
      ],
    },
    // ── Sheet 4: scenarios + grouped outline ──
    {
      name: "Scenarios",
      rows: [
        { cells: [{ value: "Rate" }, { value: 0.035 }] },
        { cells: [{ value: "Term" }, { value: 15 }] },
      ],
      scenarios: {
        current: 0,
        show: 0,
        scenarios: [
          {
            name: "High rate",
            count: 1,
            inputCells: [
              { reference: "B1", val: "0.06" },
              { reference: "B2", val: "30" },
            ],
          },
        ],
      },
    },
  ],
  chartsheets: [
    {
      name: "ChartSheet1",
      tabColor: "FF4472C4",
      chart: {
        type: "line",
        title: "Chart sheet series",
        categories: ["A", "B", "C"],
        series: [{ name: "S1", values: [1, 2, 3] }],
      },
    },
  ],
});

mkdirSync(".temp", { recursive: true });
writeFileSync(".temp/31-max-coverage.xlsx", buffer);
console.log("written 31-max-coverage.xlsx");
