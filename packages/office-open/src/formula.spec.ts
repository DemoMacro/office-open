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

  it("returns cached used dimensions per worksheet", () => {
    const options: WorkbookOptions = {
      worksheets: [
        { rows: [{ cells: [cell("C8", 1), cell("A2", 2)] }] },
        { rows: [{ cells: [cell("B4", 3)] }] },
        {},
      ],
    };
    const context = createWorkbookFormulaContext(options);

    expect(context.getRowCount(0)).toBe(8);
    expect(context.getColumnCount(0)).toBe(3);
    expect(context.getRowCount(1)).toBe(4);
    expect(context.getColumnCount(1)).toBe(2);
    expect(context.getRowCount(2)).toBe(1_048_576);
    expect(context.getColumnCount(2)).toBe(16_384);
  });

  it("clears stale cached values when recalculation yields an error or empty result", () => {
    const options: WorkbookOptions = {
      worksheets: [
        {
          rows: [
            { cells: [{ reference: "A1", value: 9, formula: "1/0" }] },
            { cells: [{ reference: "A2", value: 9, formula: "NA()" }] },
          ],
        },
      ],
    };

    calculateWorkbookFormulas(options);

    const [errorCell, emptyCell] =
      options.worksheets?.[0]?.rows?.map((row) => row.cells?.[0]) ?? [];
    expect(errorCell).toMatchObject({ error: "#DIV/0!" });
    expect(errorCell).not.toHaveProperty("value");
    expect(emptyCell).not.toHaveProperty("value");
  });

  it("evaluates a formula without parsing unrelated invalid formulas", () => {
    const options: WorkbookOptions = {
      worksheets: [
        {
          rows: [{ cells: [cell("A1", undefined, "INVALID(")] }, { cells: [cell("B1", 3)] }],
        },
      ],
    };

    expect(evaluateWorkbookFormula(options, "B1*2")).toBe(6);
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

  it("uses implicit cell addresses when calculating formulas", () => {
    const options: WorkbookOptions = {
      worksheets: [
        {
          rows: [
            { cells: [{ value: 3 }] },
            { rowNumber: 4, cells: [{ formula: "A1*2" }, { formula: "A4+1" }] },
          ],
        },
      ],
    };

    const { failures } = calculateWorkbookFormulas(options);

    expect(failures).toEqual([]);
    expect(options.worksheets?.[0]?.rows?.[1]?.cells?.map((cell) => cell.value)).toEqual([6, 7]);
  });

  it("uses implicit row numbers when evaluating hidden-row formulas", () => {
    const options: WorkbookOptions = {
      worksheets: [
        {
          rows: [
            { cells: [{ value: 2 }] },
            { hidden: true, cells: [{ value: 3 }] },
            { cells: [{ value: 5 }] },
            { cells: [{ formula: "SUBTOTAL(109,A1:A3)" }] },
          ],
        },
      ],
    };

    calculateWorkbookFormulas(options);

    expect(options.worksheets?.[0]?.rows?.[3]?.cells?.[0]?.value).toBe(7);
  });

  it("orders a local defined name ahead of a same-named global name", () => {
    const options: WorkbookOptions = {
      worksheets: [
        {
          name: "Data",
          rows: [
            { cells: [cell("C1", undefined, "RATE")] },
            { cells: [cell("B1", undefined, "A1*2")] },
            { cells: [cell("A1", 3)] },
          ],
        },
      ],
      definedNames: [
        { name: "RATE", value: "Data!A1*10" },
        { name: "RATE", value: "Data!B1*10", localSheetId: 0 },
      ],
    };

    const { failures, results } = calculateWorkbookFormulas(options);
    expect(failures).toEqual([]);
    expect(options.worksheets?.[0]?.rows?.[1]?.cells?.[0]?.value).toBe(6);
    expect(options.worksheets?.[0]?.rows?.[0]?.cells?.[0]?.value).toBe(60);
    expect(results.at(-1)?.value).toBe(60);
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

  it("distributes array caches by cell coordinates across sparse ranges", () => {
    const options: WorkbookOptions = {
      worksheets: [
        {
          rows: [
            {
              cells: [
                {
                  reference: "A1",
                  formula: { formula: "{1,2;3,4}", type: "array", reference: "A1:B2" },
                },
              ],
            },
            {
              cells: [
                { reference: "B2", formula: { formula: "", type: "array", reference: "A1:B2" } },
              ],
            },
          ],
        },
      ],
    };

    const { failures } = calculateWorkbookFormulas(options);

    expect(failures).toEqual([]);
    expect(options.worksheets?.[0]?.rows?.[0]?.cells?.[0]?.value).toBe(1);
    expect(options.worksheets?.[0]?.rows?.[1]?.cells?.[0]?.value).toBe(4);
  });

  it("reports cyclic formula dependencies and blocks their dependents", () => {
    const options: WorkbookOptions = {
      worksheets: [
        {
          rows: [
            { cells: [cell("A1", undefined, "B1+1")] },
            { cells: [cell("B1", undefined, "A1+1")] },
            { cells: [cell("C1", undefined, "A1+1")] },
            { cells: [cell("D1", undefined, "2+2")] },
          ],
        },
      ],
    };

    const { failures, results } = calculateWorkbookFormulas(options);

    expect(failures.map(({ reference }) => reference)).toEqual(["A1", "B1", "C1"]);
    expect(results.map(({ reference, value }) => [reference, value])).toEqual([["D1", 4]]);
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
