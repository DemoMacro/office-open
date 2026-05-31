import { writeFileSync } from "node:fs";

import { Workbook, Packer } from "@office-open/xlsx";

const wb = new Workbook({
  worksheets: [
    {
      name: "Validation",
      rows: [
        { cells: [{ value: "Status" }, { value: "Score" }] },
        { cells: [{ value: "Yes" }, { value: 85 }] },
        { cells: [{ value: "No" }, { value: 150 }] },
        { cells: [{ value: "Yes" }, { value: 42 }] },
        { cells: [{ value: "No" }, { value: 200 }] },
      ],
      // Data validation: dropdown list for A2:A5
      dataValidations: [
        {
          sqref: "A2:A5",
          type: "list",
          formula1: '"Yes,No,Maybe"',
          allowBlank: true,
          showErrorMessage: true,
          errorTitle: "Invalid Input",
          error: "Please select Yes, No, or Maybe",
        },
        {
          sqref: "B2:B5",
          type: "whole",
          operator: "between",
          formula1: "0",
          formula2: "100",
          showErrorMessage: true,
          errorTitle: "Invalid Score",
          error: "Score must be between 0 and 100",
        },
      ],
      // Conditional formatting: highlight scores > 100
      conditionalFormats: [
        {
          sqref: "B2:B5",
          rules: [
            {
              type: "cellIs",
              operator: "greaterThan",
              formulas: ["100"],
            },
          ],
        },
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(wb);
writeFileSync("My Workbook.xlsx", buffer);
