import { writeFileSync } from "node:fs";

import { generateWorkbook } from "@office-open/xlsx";

const buffer = await generateWorkbook({
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
              { reference: "B1", val: "0.035" },
              { reference: "B2", val: "15" },
              { reference: "B3", val: "200000" },
            ],
          },
          {
            name: "High Rate",
            count: 2,
            inputCells: [
              { reference: "B1", val: "0.075" },
              { reference: "B2", val: "30" },
              { reference: "B3", val: "350000" },
            ],
          },
          {
            name: "Short Term",
            count: 3,
            user: "Analyst",
            comment: "Aggressive payoff scenario",
            inputCells: [
              { reference: "B1", val: "0.05" },
              { reference: "B2", val: "10" },
              { reference: "B3", val: "250000" },
            ],
          },
        ],
      },
    },
  ],
});

writeFileSync("My Workbook.xlsx", buffer);
