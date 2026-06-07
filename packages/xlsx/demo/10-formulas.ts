import { writeFileSync } from "node:fs";

import { generate } from "@office-open/xlsx";

const buffer = await generate({
  title: "Formula Demo",
  worksheets: [
    {
      name: "Normal & Shared",
      rows: [
        // Row 1: Headers
        {
          cells: [
            { value: "Item", style: { font: { bold: true } } },
            { value: "Qty", style: { font: { bold: true } } },
            { value: "Price", style: { font: { bold: true } } },
            { value: "Total", style: { font: { bold: true } } },
          ],
        },
        // Row 2-4: Normal formulas (Qty * Price)
        {
          cells: [
            { value: "Apples" },
            { value: 10 },
            { value: 2.5 },
            { formula: { formula: "B2*C2" }, value: 25 },
          ],
        },
        {
          cells: [
            { value: "Bananas" },
            { value: 5 },
            { value: 3.0 },
            { formula: { formula: "B3*C3" }, value: 15 },
          ],
        },
        {
          cells: [
            { value: "Cherries" },
            { value: 8 },
            { value: 4.5 },
            { formula: { formula: "B4*C4" }, value: 36 },
          ],
        },
        // Row 5: SUM formulas without cached results (Excel recalculates on open)
        {
          cells: [
            { value: "Total", style: { font: { bold: true } } },
            { formula: { formula: "SUM(B2:B4)" } },
            { formula: { formula: "SUM(C2:C4)" } },
            { formula: { formula: "SUM(D2:D4)" } },
          ],
        },
        // Row 7-9: Shared formula example
        // Definition cell: t="shared", ref=full range, si=group index
        {
          rowNumber: 7,
          cells: [
            { value: "X" },
            { value: "Y" },
            { value: "X+Y", style: { font: { bold: true } } },
          ],
        },
        {
          rowNumber: 8,
          cells: [
            { value: 1 },
            { value: 2 },
            {
              reference: "C8",
              formula: { formula: "A8+B8", type: "shared", reference: "C8:C10", sharedIndex: 0 },
              value: 3,
            },
          ],
        },
        // Follower cells: t="shared", si=same index, NO ref, NO formula text
        {
          rowNumber: 9,
          cells: [
            { value: 4 },
            { value: 5 },
            {
              reference: "C9",
              formula: { formula: "", type: "shared", sharedIndex: 0 },
              value: 9,
            },
          ],
        },
        {
          rowNumber: 10,
          cells: [
            { value: 6 },
            { value: 7 },
            {
              reference: "C10",
              formula: { formula: "", type: "shared", sharedIndex: 0 },
              value: 13,
            },
          ],
        },
      ],
    },
    {
      name: "Array Formula",
      rows: [
        // Row 1: Headers
        {
          cells: [{ value: "X" }, { value: "Y" }, { value: "X*Y (array)" }],
        },
        // Row 2-4: Source data
        { cells: [{ value: 1 }, { value: 10 }] },
        { cells: [{ value: 2 }, { value: 20 }] },
        { cells: [{ value: 3 }, { value: 30 }] },
        // Row 5-7: Array formula range (A5:B7 = A2:A4*B2:B4)
        // Definition cell: t="array", ref=full range, formula text present
        {
          rowNumber: 5,
          cells: [
            {
              reference: "A5",
              formula: { formula: "A2:A4*B2:B4", type: "array", reference: "A5:B7" },
              value: 10,
            },
            {
              reference: "B5",
              formula: { formula: "", type: "array", reference: "A5:B7" },
              value: 100,
            },
          ],
        },
        // Follower cells: t="array", ref=full range, NO formula text
        {
          rowNumber: 6,
          cells: [
            {
              reference: "A6",
              formula: { formula: "", type: "array", reference: "A5:B7" },
              value: 20,
            },
            {
              reference: "B6",
              formula: { formula: "", type: "array", reference: "A5:B7" },
              value: 200,
            },
          ],
        },
        {
          rowNumber: 7,
          cells: [
            {
              reference: "A7",
              formula: { formula: "", type: "array", reference: "A5:B7" },
              value: 30,
            },
            {
              reference: "B7",
              formula: { formula: "", type: "array", reference: "A5:B7" },
              value: 300,
            },
          ],
        },
      ],
    },
  ],
});

writeFileSync("My Workbook.xlsx", buffer);
