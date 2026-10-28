import type { WorkbookOptions } from "@office-open/xlsx";
import { describe, expect, it } from "vite-plus/test";

import {
  calculateWorkbookFormulas,
  createWorkbookFormulaContext,
  evaluateWorkbookFormula,
} from "./formula";
import { generateWorkbook, parseWorkbook } from "./workbook";

function cell(reference: string, value?: string | number, formula?: string) {
  return {
    reference,
    ...(value === undefined ? {} : { value }),
    ...(formula === undefined ? {} : { formula }),
  };
}

describe("canonical workbook formula integration", () => {
  it("evaluates formulas in dependency order", () => {
    const options: WorkbookOptions = {
      worksheets: [
        {
          name: "Data",
          rows: [
            { cells: [cell("A1", 2)] },
            { cells: [cell("B1", undefined, "A1*2")] },
            { cells: [cell("C1", undefined, "B1+RATE")] },
          ],
        },
      ],
      definedNames: [{ name: "RATE", value: "Data!A1*10" }],
    };
    const { results, failures } = calculateWorkbookFormulas(options);
    expect(failures).toEqual([]);
    expect(results.map((result) => result.value)).toEqual([4, 24]);
    expect(options.worksheets?.[0]?.rows?.[1]?.cells?.[0]?.value).toBe(4);
    expect(options.worksheets?.[0]?.rows?.[2]?.cells?.[0]?.value).toBe(24);
  });

  it("orders formulas through defined-name dependencies", () => {
    const options: WorkbookOptions = {
      worksheets: [
        {
          name: "Data",
          rows: [
            { cells: [cell("A1", 2)] },
            { cells: [cell("C1", undefined, "RATE")] },
            { cells: [cell("B1", undefined, "A1*2")] },
          ],
        },
      ],
      definedNames: [{ name: "RATE", value: "Data!B1*10" }],
    };
    const { failures, results } = calculateWorkbookFormulas(options);
    expect(failures).toEqual([]);
    expect(options.worksheets?.[0]?.rows?.at(-1)?.cells?.[0]?.value).toBe(4);
    expect(results.at(-1)?.value).toBe(40);
  });

  it("rewrites shared formulas and distributes array caches", () => {
    const options: WorkbookOptions = {
      worksheets: [
        {
          rows: [
            { cells: [cell("A2", 1), cell("B2", 2)] },
            { cells: [cell("A3", 3), cell("B3", 4)] },
            {
              cells: [
                {
                  reference: "C2",
                  formula: { formula: "A2+B2", type: "shared", reference: "C2:C3", sharedIndex: 0 },
                },
                {
                  reference: "C3",
                  formula: { formula: "", type: "shared", reference: "C2:C3", sharedIndex: 0 },
                },
                {
                  reference: "D2",
                  formula: { formula: "A2:A3*B2:B3", type: "array", reference: "D2:D3" },
                },
                {
                  reference: "D3",
                  formula: { formula: "", type: "array", reference: "D2:D3" },
                },
              ],
            },
          ],
        },
      ],
    };
    const { failures } = calculateWorkbookFormulas(options);
    expect(failures).toEqual([]);
    const cells = options.worksheets?.[0]?.rows?.at(-1)?.cells;
    expect(cells?.[0]?.value).toBe(3);
    expect(cells?.[1]?.value).toBe(7);
    expect(cells?.[2]?.value).toBe(2);
    expect(cells?.[3]?.value).toBe(12);
  });

  it("uses tables, hidden-row filters, and structured errors", () => {
    const options: WorkbookOptions = {
      worksheets: [
        {
          name: "Sales",
          rows: [
            { cells: [cell("A1", "Alpha"), cell("B1", "Beta")] },
            { hidden: true, cells: [cell("A2", 10), cell("B2", 1)] },
            { cells: [cell("A3", 20), cell("B3", "#DIV/0!")] },
            { cells: [cell("B5", undefined, "SUBTOTAL(109,Sales[Alpha])")] },
            { cells: [cell("B6", undefined, "B5/B4")] },
          ],
          tables: [
            {
              displayName: "Sales",
              name: "Sales",
              ref: "A1:B3",
              columns: [{ name: "Alpha" }, { name: "Beta" }],
            },
          ],
        },
      ],
    };
    const context = createWorkbookFormulaContext(options);
    expect(context.getTableName("sales")?.columnNames).toEqual(["Alpha", "Beta"]);
    const { results, failures } = calculateWorkbookFormulas(options);
    expect(failures).toEqual([]);
    expect(results.at(-2)?.value).toBe(20);
    expect(results.at(-1)?.value).toBe("#DIV/0!");
    expect(options.worksheets?.[0]?.rows?.at(-1)?.cells?.[0]?.error).toBe("#DIV/0!");
  });

  it("evaluates OpenFormula cells for ODS workbooks", () => {
    const options: WorkbookOptions = {
      worksheets: [
        {
          name: "Sheet1",
          rows: [
            { cells: [cell("A1", 2), cell("A2", 3)] },
            { cells: [cell("B1", undefined, "[.A1]+[.A2]")] },
          ],
        },
      ],
    };
    expect(
      evaluateWorkbookFormula(
        options,
        "[.A1]+[.A2]",
        { sheet: 0, row: 2, column: 2 },
        {
          dialect: "openformula",
        },
      ),
    ).toBe(5);
    const { failures, results } = calculateWorkbookFormulas(options, { dialect: "openformula" });
    expect(failures).toEqual([]);
    expect(results[0]?.value).toBe(5);
  });

  it("normalizes ODF named-range addresses for OpenFormula", () => {
    const options: WorkbookOptions = {
      worksheets: [
        {
          name: "Sheet1",
          rows: [
            { cells: [cell("A1", 2), cell("B1", 3)] },
            { cells: [cell("C1", undefined, "SUM(Values)")] },
          ],
        },
      ],
      definedNames: [{ name: "Values", value: "$Sheet1.A1:B1", localSheetId: 0 }],
    };
    const { failures, results } = calculateWorkbookFormulas(options, { dialect: "openformula" });
    expect(failures).toEqual([]);
    expect(results[0]?.value).toBe(5);
  });

  it("recalculates formulas after ODS package round-trip", async () => {
    const options: WorkbookOptions = {
      worksheets: [
        {
          name: "Sheet1",
          rows: [
            { cells: [cell("A1", 2), cell("B1", 3)] },
            { cells: [cell("C1", undefined, "[.A1]+[.B1]")] },
          ],
        },
      ],
      definedNames: [{ name: "Values", value: "$Sheet1.A1:B1", localSheetId: 0 }],
    };
    const parsed = await parseWorkbook(await generateWorkbook("ods", options, "uint8array"));
    const { failures, results } = calculateWorkbookFormulas(parsed, { dialect: "openformula" });
    expect(failures).toEqual([]);
    expect(results[0]?.value).toBe(5);
  });
});
