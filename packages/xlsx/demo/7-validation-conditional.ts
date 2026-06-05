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
      // Data validation: disable input prompts for this sheet
      dataValidationsDisablePrompts: true,
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
      // Conditional formatting: cellIs rule for scores > 100
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
    {
      name: "ColorScale",
      rows: [
        { cells: [{ value: "Temperature" }, { value: "Humidity" }, { value: "Pressure" }] },
        { cells: [{ value: 10 }, { value: 20 }, { value: 95 }] },
        { cells: [{ value: 50 }, { value: 60 }, { value: 70 }] },
        { cells: [{ value: 90 }, { value: 80 }, { value: 45 }] },
        { cells: [{ value: 30 }, { value: 40 }, { value: 85 }] },
        { cells: [{ value: 70 }, { value: 10 }, { value: 30 }] },
      ],
      conditionalFormats: [
        // Color scale: 3-color gradient
        {
          sqref: "A2:A6",
          rules: [
            {
              type: "colorScale",
              priority: 1,
              colorScale: {
                cfvo: [{ type: "min" }, { type: "percentile", val: 50 }, { type: "max" }],
                colors: ["F8696B", "FFEB84", "63BE7B"],
              },
            },
          ],
        },
        // Data bar
        {
          sqref: "B2:B6",
          rules: [
            {
              type: "dataBar",
              priority: 2,
              dataBar: {
                cfvo: [{ type: "min" }, { type: "max" }],
                color: "638EC6",
                showValue: true,
              },
            },
          ],
        },
        // Icon set
        {
          sqref: "C2:C6",
          rules: [
            {
              type: "iconSet",
              priority: 3,
              iconSet: {
                cfvo: [
                  { type: "percent", val: 0 },
                  { type: "percent", val: 33 },
                  { type: "percent", val: 67 },
                ],
                iconSet: "3TrafficLights1",
                showValue: true,
              },
            },
          ],
        },
      ],
    },
    {
      name: "DataBarNoValue",
      rows: [
        { cells: [{ value: "Progress" }] },
        { cells: [{ value: 25 }] },
        { cells: [{ value: 50 }] },
        { cells: [{ value: 75 }] },
        { cells: [{ value: 100 }] },
      ],
      conditionalFormats: [
        {
          sqref: "A2:A5",
          rules: [
            {
              type: "dataBar",
              dataBar: {
                cfvo: [
                  { type: "num", val: 0 },
                  { type: "num", val: 100 },
                ],
                color: "00B050",
                showValue: false,
              },
            },
          ],
        },
      ],
    },
  ],
});

const buffer = await Packer.toBuffer(wb);
writeFileSync("My Workbook.xlsx", buffer);
