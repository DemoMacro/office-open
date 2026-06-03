import { writeFileSync } from "node:fs";

import { Workbook, Packer } from "@office-open/xlsx";

const wb = new Workbook({
  worksheets: [
    {
      name: "Loan Calculator",
      rows: [
        {
          cells: [{ value: "Interest Rate" }, { value: 0.05, style: { numFmt: "0.00%" } }],
        },
        {
          cells: [{ value: "Term (years)" }, { value: 30 }],
        },
        {
          cells: [{ value: "Loan Amount" }, { value: 250000, style: { numFmt: "#,##0" } }],
        },
        {
          cells: [
            { value: "Monthly Payment" },
            {
              formula: { formula: "PMT(B1/12,B2*12,-B3)" },
              style: { numFmt: "#,##0.00" },
            },
          ],
        },
      ],
      scenarios: {
        current: 0,
        show: 0,
        scenarios: [
          {
            name: "Low Rate",
            count: 1,
            inputCells: [
              { r: "B1", val: "0.035" },
              { r: "B2", val: "15" },
              { r: "B3", val: "200000" },
            ],
          },
          {
            name: "High Rate",
            count: 2,
            inputCells: [
              { r: "B1", val: "0.075" },
              { r: "B2", val: "30" },
              { r: "B3", val: "350000" },
            ],
          },
          {
            name: "Short Term",
            count: 3,
            user: "Analyst",
            comment: "Aggressive payoff scenario",
            inputCells: [
              { r: "B1", val: "0.05" },
              { r: "B2", val: "10" },
              { r: "B3", val: "250000" },
            ],
          },
        ],
      },
    },
  ],
});

const buffer = await Packer.toBuffer(wb);
writeFileSync("My Workbook.xlsx", buffer);
