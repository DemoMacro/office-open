import { describe, expect, it } from "vite-plus/test";

import type { WorkbookOptions } from "../parts/file";
import { lintWorkbookFormulas } from "./formula-lint";

/** Workbook with one Data sheet whose cells carry the given formulas, plus
 *  optional extra named (empty) sheets. */
const workbookWithFormulas = (
  formulas: Record<string, string | undefined>,
  extraSheetNames: string[] = [],
): WorkbookOptions => {
  const rows = Object.entries(formulas).map(([reference, formula]) => ({
    cells: [{ reference, formula }],
  }));
  const worksheets: WorkbookOptions["worksheets"] = [{ name: "Data", rows }];
  for (const name of extraSheetNames) worksheets.push({ name });
  return { worksheets };
};

describe("lintWorkbookFormulas", () => {
  it("passes local references and known sheet names", () => {
    const options = workbookWithFormulas(
      {
        B1: "SUM(A1:A2)",
        B2: "Data!A1",
        B3: "'My Sheet'!A1",
        B4: "SUM(Data!A1:Data!B2)",
      },
      ["My Sheet"],
    );
    expect(lintWorkbookFormulas(options)).toEqual([]);
  });

  it("resolves default sheet names (Sheet1, Sheet2, …)", () => {
    const options: WorkbookOptions = {
      worksheets: [
        { rows: [{ cells: [{ reference: "A1", formula: "Sheet2!B1" }] }] },
        { rows: [] },
      ],
    };
    expect(lintWorkbookFormulas(options)).toEqual([]);
  });

  it("flags references to sheets that do not exist", () => {
    const issues = lintWorkbookFormulas(workbookWithFormulas({ B1: "SUM(Sheet9!A1:A2)" }));
    expect(issues).toHaveLength(1);
    expect(issues[0]!.location).toBe("Data!B1");
    expect(issues[0]!.message).toContain('"Sheet9"');
    expect(issues[0]!.message).toContain("sheets: Data");
  });

  it("checks each leg of a 3-D range and quoted names", () => {
    const issues = lintWorkbookFormulas(
      workbookWithFormulas({ B1: "SUM(North:Data!A1)", B2: "'Bad Sheet'!A1" }),
    );
    expect(issues.map((i) => i.message)).toEqual([
      expect.stringContaining('"North"'),
      expect.stringContaining('"Bad Sheet"'),
    ]);
  });

  it("does not read the range half of Sheet!A1:Sheet!B2 as a sheet name", () => {
    // The scan back from the second '!' collects "A1:Data" — the A1 half is a
    // cell reference, not a missing sheet.
    expect(lintWorkbookFormulas(workbookWithFormulas({ B4: "SUM(Data!A1:Data!B2)" }))).toEqual([]);
  });

  it("does not misread string literals, error literals, or external refs", () => {
    const options = workbookWithFormulas({
      B1: '=IF(A1="gotcha!", "ok!!", "no")',
      B2: "#REF!+A1",
      B3: "[1]Sheet9!A1",
      B4: 'SUM(A1:A2)&"!"',
    });
    expect(lintWorkbookFormulas(options)).toEqual([]);
  });

  it("flags unbalanced parentheses and odd quotes", () => {
    const issues = lintWorkbookFormulas(
      workbookWithFormulas({ B1: "SUM(A1:A2", B2: '="unclosed' }),
    );
    expect(issues.map((i) => i.message)).toEqual([
      "unbalanced parentheses",
      "unbalanced double quotes",
    ]);
  });

  it("lints conditional-format and data-validation formulas", () => {
    const options: WorkbookOptions = {
      worksheets: [
        {
          name: "Data",
          conditionalFormats: [
            { sqref: "A1:A10", rules: [{ type: "expression", formulas: ["Missing!A1>1"] }] },
          ],
          dataValidations: [{ sqref: "B1:B10", type: "list", formula1: "Missing2!$A$1:$A$5" }],
        },
      ],
    };
    const issues = lintWorkbookFormulas(options);
    expect(issues.map((i) => i.location)).toEqual([
      "Data conditional format A1:A10",
      "Data data validation B1:B10",
    ]);
  });
});
