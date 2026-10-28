import { describe, expect, it } from "vitest";

import type { FormulaReferenceArea } from "./ast";
import { formulaFunctions } from "./catalog";
import { createFormulaContext, type FormulaWorkbookContext } from "./context";
import { FormulaValueError } from "./errors";
import { evaluateFormula, evaluateFormulaAst } from "./evaluator";
import { getFunction, getFunctionMetadata, listFunctionNames } from "./functions";
import {
  buildDependencyGraph,
  collectDependencies,
  collectDirtyFormulas,
  detectFormulaCycles,
  formatFormulaDependencyTarget,
  topologicalFormulaOrder,
} from "./graph";
import * as formulaPublic from "./index";
import { tokenizeFormula } from "./lexer";
import { parseFormula } from "./parser";
import { formatReferenceArea } from "./references";
import { shiftFormula, shiftFormulaAst } from "./rewrite";
import { resolveSharedFormula } from "./shared";
import { stringifyFormula } from "./stringify";
import { FormulaError } from "./value";

describe("formula tokenizer", () => {
  it("tokenizes literals, references, and calls", () => {
    const tokens = tokenizeFormula('=SUM(A1:B2) & "x"', "excel").filter(
      (token) => token.kind !== "whitespace",
    );
    expect(tokens.map((token) => token.kind)).toEqual([
      "operator",
      "identifier",
      "open-paren",
      "identifier",
      "operator",
      "identifier",
      "close-paren",
      "operator",
      "string",
      "end",
    ]);
  });
});

describe("formula parser", () => {
  it("parses and preserves calls, ranges, operators, and literals", () => {
    const ast = parseFormula('=SUM(A1:B2, 2, "x")');
    expect(ast).toMatchObject({
      kind: "call",
      name: "SUM",
      args: [{ kind: "reference" }, { kind: "number", value: 2 }, { kind: "string", value: "x" }],
    });
    expect(ast.raw).toBe('=SUM(A1:B2, 2, "x")');
  });

  it("parses quoted sheet references", () => {
    const ast = parseFormula("='My Sheet'!$B$2");
    expect(ast.kind).toBe("reference");
    expect(ast.raw).toBe("='My Sheet'!$B$2");
  });

  it("parses spill references", () => {
    expect(parseFormula("=A1#")).toMatchObject({
      kind: "reference",
      reference: { spill: true },
    });
    expect(() => parseFormula("=Sheet1:Sheet2!A1#")).toThrow();
  });

  it("parses quoted 3D sheet references", () => {
    const ast = parseFormula("='First Sheet':'Second Sheet'!$B$2");
    expect(ast).toMatchObject({
      kind: "reference",
      reference: {
        sheets: { names: ["First Sheet", "Second Sheet"], range: true },
        areas: [
          {
            fromSheet: "First Sheet",
            toSheet: "Second Sheet",
            start: { row: 2, column: 2 },
          },
        ],
      },
    });
  });

  it("parses reference set operators", () => {
    expect(parseFormula("=A1:B2 A2:B3")).toMatchObject({
      kind: "binary",
      operator: " ",
    });
    expect(parseFormula("=[.A1]~[.B2]", { dialect: "openformula" })).toMatchObject({
      kind: "binary",
      operator: "~",
    });
    expect(parseFormula("=[.A1]![.B2]", { dialect: "openformula" })).toMatchObject({
      kind: "binary",
      operator: "!",
    });
  });

  it("parses structured whole-column references", () => {
    const ast = parseFormula("=A:C");
    expect(ast).toMatchObject({
      kind: "reference",
      reference: {
        areas: [
          {
            start: { kind: "whole-column", column: 1 },
            end: { kind: "whole-column", column: 3 },
          },
        ],
      },
    });
  });

  it("parses whole-row references and rejects mixed reference axes", () => {
    expect(parseFormula("=5:7")).toMatchObject({
      kind: "reference",
      reference: {
        areas: [
          {
            start: { kind: "whole-row", row: 5 },
            end: { kind: "whole-row", row: 7 },
          },
        ],
      },
    });
    expect(parseFormula("=$5:$7")).toMatchObject({
      kind: "reference",
      reference: {
        areas: [
          {
            start: { kind: "whole-row", row: 5, absolute: true },
            end: { kind: "whole-row", row: 7, absolute: true },
          },
        ],
      },
    });
    expect(() => parseFormula("=A1:5")).toThrow();
    expect(() => parseFormula("=A:A2")).toThrow();
    expect(() => parseFormula("=A:C1")).toThrow();
    expect(parseFormula("=A1:B2:C3")).toMatchObject({
      kind: "reference",
      reference: {
        areas: [
          {
            start: { row: 1, column: 1 },
            end: { row: 3, column: 3 },
          },
        ],
      },
    });
  });

  it("formats whole-axis references without duplicating either edge", () => {
    expect(stringifyFormula(parseFormula("=SUM(A:C)"))).toBe("=SUM(A:C)");
    expect(stringifyFormula(parseFormula("=SUM($5:$7)"))).toBe("=SUM($5:$7)");
    expect(stringifyFormula(parseFormula("=Sheet1:Sheet2!5:7"))).toBe("=Sheet1:Sheet2!5:7");
    expect(stringifyFormula(parseFormula("=[Book]Sheet1:Sheet2!5:7"))).toBe(
      "=[Book]Sheet1:Sheet2!5:7",
    );
    expect(shiftFormula("=SUM(A:C)", { rows: 1, columns: 1 })).toBe("=SUM(B:D)");
    expect(shiftFormula("=SUM(5:7)", { rows: 1, columns: 1 })).toBe("=SUM(6:8)");
    expect(shiftFormula("=SUM($5:$7)", { rows: 1, columns: 1 })).toBe("=SUM($5:$7)");
  });

  it("parses R1C1 text as a name instead of truncating it", () => {
    expect(parseFormula("=R1C1")).toMatchObject({ kind: "name", name: "R1C1" });
    expect(stringifyFormula(parseFormula("=R1C1"))).toBe("=R1C1");
    expect(stringifyFormula(parseFormula("=SUM(R1C1)"))).toBe("=SUM(R1C1)");
  });

  it("round-trips implicit intersection on references and tables", () => {
    expect(stringifyFormula(parseFormula("=(@A1)"))).toBe("=(@A1)");
    expect(stringifyFormula(parseFormula("=@Sheet1!A1"))).toBe("=@Sheet1!A1");
    expect(stringifyFormula(parseFormula("=@Sales[Amount]"))).toBe("=@Sales[Amount]");
    expect(parseFormula("=Sales[[Name]:[Amount]]")).toMatchObject({
      reference: {
        table: {
          selectors: [{ kind: "range", start: "Name", end: "Amount" }],
        },
      },
    });
    expect(stringifyFormula(parseFormula("=Sales[[Name]:[Amount]]"))).toBe(
      "=Sales[[Name]:[Amount]]",
    );
    expect(stringifyFormula(parseFormula("=Sales[@Name]"))).toBe("=Sales[[#This Row],[Name]]");
    expect(() => parseFormula("=Sales[[]]")).toThrow();
    expect(() => parseFormula("=Sales[[Alpha]:[]]")).toThrow();
    expect(() => parseFormula("=Sales[Alpha]#")).toThrow();
    expect(parseFormula("=Sales['#Amount]")).toMatchObject({
      reference: { table: { selectors: [{ kind: "column", column: "#Amount" }] } },
    });
    expect(stringifyFormula(parseFormula("=Sales['#Amount]"))).toBe("=Sales['#Amount]");
    expect(parseFormula("=Sales['[Temp']]")).toMatchObject({
      reference: { table: { selectors: [{ kind: "column", column: "[Temp]" }] } },
    });
    expect(stringifyFormula(parseFormula("=Sales['[Temp']]"))).toBe("=Sales['[Temp']]");
    expect(parseFormula("=Sales[Temp']]")).toMatchObject({
      reference: { table: { selectors: [{ kind: "column", column: "Temp]" }] } },
    });
    expect(stringifyFormula(parseFormula("=Sales[Temp']]"))).toBe("=Sales[Temp']]");
    expect(parseFormula("=Sales[']Temp]")).toMatchObject({
      reference: { table: { selectors: [{ kind: "column", column: "]Temp" }] } },
    });
    expect(stringifyFormula(parseFormula("=Sales[']Temp]"))).toBe("=Sales[']Temp]");
    expect(parseFormula("=Sales[A''B]")).toMatchObject({
      reference: { table: { selectors: [{ kind: "column", column: "A'B" }] } },
    });
    expect(() => parseFormula("=Sales[A'B]")).toThrow();
    expect(() => parseFormula("=Sales['Amount]")).toThrow();
    expect(() => parseFormula("=Sales[]")).toThrow();
    expect(() => parseFormula("=Sales[#Bogus]")).toThrow();
    expect(() => parseFormula("=Sales[,]")).toThrow();
    expect(() => parseFormula("=Sales[[#All][Amount]]")).toThrow();
  });

  it("formats workbook identities across formula dialects", () => {
    const cell: FormulaReferenceArea["start"] = {
      row: 1,
      column: 1,
      absoluteRow: false,
      absoluteColumn: false,
    };
    const area = (
      workbook: NonNullable<FormulaReferenceArea["workbook"]>,
    ): FormulaReferenceArea => ({
      workbook,
      start: cell,
    });
    expect(formatReferenceArea(area({ path: "file.xlsx" }))).toBe("[file.xlsx]A1");
    expect(formatReferenceArea(area({ index: 2 }))).toBe("[2]A1");
    expect(formatReferenceArea(area({ name: "Book" }), "openformula")).toBe("['Book'#.A1]");
    const sheetArea: FormulaReferenceArea = {
      workbook: { name: "Book" },
      fromSheet: "Sheet",
      start: cell,
    };
    expect(formatReferenceArea(sheetArea, "openformula")).toBe("['Book'#Sheet.A1]");
    expect(stringifyFormula(parseFormula("=['Book'#Sheet.A1]", { dialect: "openformula" }))).toBe(
      "=['Book'#Sheet.A1]",
    );
    expect(stringifyFormula(parseFormula("=['It''s'.A1]", { dialect: "openformula" }))).toBe(
      "=['It''s'.A1]",
    );
    expect(stringifyFormula(parseFormula("=['B]ook']Sheet!A1"))).toBe("=['B]ook']Sheet!A1");
  });

  it("round-trips OpenFormula whole-axis references", () => {
    expect(stringifyFormula(parseFormula("=[Sheet.$A:Sheet.$C]", { dialect: "openformula" }))).toBe(
      "=[Sheet.$A:Sheet.$C]",
    );
    expect(stringifyFormula(parseFormula("=[Sheet.$5:Sheet.$7]", { dialect: "openformula" }))).toBe(
      "=[Sheet.$5:Sheet.$7]",
    );
    expect(parseFormula("=[Sheet.$A:Sheet.$C]", { dialect: "openformula" })).toMatchObject({
      kind: "reference",
      reference: {
        sheets: { names: ["Sheet"], range: false },
        areas: [
          {
            fromSheet: "Sheet",
            start: { kind: "whole-column", column: 1, absolute: true },
            end: { kind: "whole-column", column: 3, absolute: true },
          },
        ],
      },
    });
  });

  it("parses OpenFormula absolute sheet locators", () => {
    expect(parseFormula("=[$Sheet.A1]", { dialect: "openformula" })).toMatchObject({
      reference: { sheets: { names: ["Sheet"], range: false } },
    });
    expect(parseFormula("=[$'Work Sheet'.A1]", { dialect: "openformula" })).toMatchObject({
      reference: { sheets: { names: ["Work Sheet"], range: false } },
    });
    expect(
      stringifyFormula(parseFormula("=[$Sheet.A1]", { dialect: "openformula" }), {
        dialect: "openformula",
      }),
    ).toBe("=[Sheet.A1]");
  });

  it("rejects OpenFormula mixed-axis references structurally", () => {
    expect(() => parseFormula("=[.A1:.5]", { dialect: "openformula" })).toThrow();
    expect(() => parseFormula("=[.A:.5]", { dialect: "openformula" })).toThrow();
    expect(() => parseFormula("=[.A1:.5:.B1]", { dialect: "openformula" })).toThrow();
  });

  it("rejects references beyond the Excel worksheet boundary", () => {
    expect(parseFormula("=XFD1048576").kind).toBe("reference");
    expect(() => parseFormula("=XFE1")).toThrow();
    expect(() => parseFormula("=A1048577")).toThrow();
    expect(parseFormula("=A:XFD").kind).toBe("reference");
    expect(() => parseFormula("=A:XFE")).toThrow();
    expect(() => parseFormula("=[Sheet.XFE]", { dialect: "openformula" })).toThrow();
    expect(() => parseFormula("=[Sheet.A1048577]", { dialect: "openformula" })).toThrow();
    expect(() => parseFormula("=[Sheet.$1048577]", { dialect: "openformula" })).toThrow();
  });

  it("preserves OpenFormula bracket references as references", () => {
    const ast = parseFormula("=SUM([.A1:.A3])", { dialect: "openformula" });
    expect(ast).toMatchObject({ kind: "call", name: "SUM" });
    expect(ast.raw).toBe("=SUM([.A1:.A3])");
    expect(stringifyFormula(ast, { dialect: "openformula" })).toBe("=SUM([.A1:.A3])");
  });

  it("rejects trailing input", () => {
    expect(() => parseFormula("=1+")).toThrow();
  });

  it("wraps lexical failures as FormulaParseError", () => {
    expect(() => parseFormula('="open')).toThrowError(/unterminated string/u);
  });
});

describe("formula analysis and rewriting", () => {
  it("collects cell, range, and named dependencies", () => {
    const ast = parseFormula("=SUM(A1:B2)+Rate+Table[Amount]");
    expect(collectDependencies(ast)).toEqual([
      {
        target: { kind: "range", from: { row: 1, column: 1 }, to: { row: 2, column: 2 } },
        sourceStart: 5,
        sourceEnd: 10,
      },
      { target: { kind: "name", name: "Rate" }, sourceStart: 12, sourceEnd: 16 },
      {
        target: {
          kind: "table",
          name: "Table",
          table: { table: "Table", selectors: [{ kind: "column", column: "Amount" }] },
        },
        sourceStart: 17,
        sourceEnd: 30,
      },
    ]);
  });

  it("round-trips formulas without inventing an equals prefix", () => {
    const ast = parseFormula("SUM(A1:B2)");
    expect(ast.prefix).toBeUndefined();
    expect(stringifyFormula(ast)).toBe("SUM(A1:B2)");
  });

  it("round-trips structured table selectors", () => {
    const ast = parseFormula("=Table[[#All],[Amount]:[Total]]");
    expect(ast.kind).toBe("reference");
    expect(stringifyFormula(ast)).toBe("=Table[[#All],[Amount]:[Total]]");
  });

  it("preserves ordered table column selectors", () => {
    const ast = parseFormula("=Table[[Gamma]:[Alpha],[Beta]]");
    expect(ast).toMatchObject({
      reference: {
        table: {
          selectors: [
            { kind: "range", start: "Gamma", end: "Alpha" },
            { kind: "column", column: "Beta" },
          ],
        },
      },
    });
    expect(stringifyFormula(ast)).toBe("=Table[[Gamma]:[Alpha],[Beta]]");
  });

  it("uses one execution registry for value and AST-aware functions", () => {
    expect(getFunctionMetadata("NOW")).toMatchObject({ minArgs: 0, maxArgs: 0 });
    expect(getFunctionMetadata("IF")).toMatchObject({ minArgs: 2, maxArgs: 3 });
    expect(getFunction("NOW")?.evaluateAst).toBeUndefined();
    expect(getFunction("IF")?.evaluateAst).toBeTypeOf("function");
  });

  it("shifts relative references without changing absolute components", () => {
    expect(shiftFormula("=A1+$A1+A$1+$A$1", { rows: 2, columns: 1 })).toBe("=B3+$A3+B$1+$A$1");
    expect(shiftFormula("=Sheet1!A1+'My Sheet'!B2:C2", { rows: 1, columns: 1 })).toBe(
      "=Sheet1!B2+'My Sheet'!C3:D3",
    );
  });

  it("rejects shifts outside the worksheet instead of clamping", () => {
    expect(() => shiftFormula("=A1", { rows: -1 })).toThrow(RangeError);
    expect(() => shiftFormula("=A1", { columns: -1 })).toThrow(RangeError);
  });

  it("builds cell dependency cycles and dirty propagation", () => {
    const nodes = [
      { id: "A1", ast: parseFormula("=B1+1"), cell: { sheet: "Sheet1", row: 1, column: 1 } },
      { id: "B1", ast: parseFormula("=A1+1"), cell: { sheet: "Sheet1", row: 1, column: 2 } },
      { id: "C1", ast: parseFormula("=A1:C1"), cell: { sheet: "Sheet1", row: 1, column: 3 } },
      { id: "D1", ast: parseFormula("=1"), cell: { sheet: "Sheet1", row: 1, column: 4 } },
    ];
    const graph = buildDependencyGraph(nodes);
    expect(detectFormulaCycles(graph)).toEqual([["A1", "B1"], ["C1"]]);
    expect(collectDirtyFormulas(graph, ["A1"])).toEqual(new Set(["A1", "B1", "C1"]));
  });

  it("builds 3D dependencies across intermediate sheets", () => {
    const nodes = [
      {
        id: "Summary",
        ast: parseFormula("=First:End!A1"),
        cell: { sheet: "First", row: 1, column: 2 },
      },
      { id: "Middle", ast: parseFormula("=1"), cell: { sheet: "Middle", row: 1, column: 1 } },
      { id: "End", ast: parseFormula("=2"), cell: { sheet: "End", row: 1, column: 1 } },
    ];
    const graph = buildDependencyGraph(nodes, { sheetOrder: ["First", "Middle", "End"] });
    expect([...graph.edges.get("Summary")!].sort()).toEqual(["End", "Middle"]);
  });

  it("keeps flat sheet order conservative across workbooks", () => {
    const nodes = [
      {
        id: "summary",
        ast: parseFormula("[Book1]First:End!A1"),
        cell: { workbook: { name: "Book1" }, sheet: "First", row: 1, column: 2 },
      },
      {
        id: "outside",
        ast: parseFormula("[Book1]Outside!A1"),
        cell: { workbook: { name: "Book1" }, sheet: "Outside", row: 1, column: 1 },
      },
      {
        id: "other-workbook",
        ast: parseFormula("[Book2]Data!A1"),
        cell: { workbook: { name: "Book2" }, sheet: "Data", row: 1, column: 1 },
      },
    ];
    const graph = buildDependencyGraph(nodes, {
      sheetOrder: ["First", "Between", "End"],
      getSheetId: (workbook, sheet) =>
        typeof workbook !== "string" && workbook?.name === "Book1"
          ? { First: 1, Outside: 2, End: 3 }[sheet]
          : undefined,
    });
    expect([...graph.edges.get("summary")!]).toEqual(["outside"]);
  });

  it("resolves 3D dependencies by sheet ids regardless of node order", () => {
    const nodes = [
      {
        id: "Summary",
        ast: parseFormula("=First:Last!A1"),
        cell: { sheet: "First", row: 1, column: 2 },
      },
      { id: "Last", ast: parseFormula("=1"), cell: { sheet: "Last", row: 1, column: 1 } },
      { id: "Middle", ast: parseFormula("=2"), cell: { sheet: "Middle", row: 1, column: 1 } },
      { id: "Outside", ast: parseFormula("=3"), cell: { sheet: "Outside", row: 1, column: 1 } },
    ];
    const graph = buildDependencyGraph(nodes, {
      getSheetId: (_workbook, sheet) => ({ First: 1, Middle: 2, Last: 3, Outside: 9 })[sheet],
    });
    expect([...graph.edges.get("Summary")!].sort()).toEqual(["Last", "Middle"]);
  });

  it("conservatively matches 3D dependencies without sheet metadata", () => {
    const nodes = [
      {
        id: "Summary",
        ast: parseFormula("=First:Last!A1"),
        cell: { sheet: "First", row: 1, column: 2 },
      },
      { id: "Last", ast: parseFormula("=1"), cell: { sheet: "Last", row: 1, column: 1 } },
      { id: "Middle", ast: parseFormula("=2"), cell: { sheet: "Middle", row: 1, column: 1 } },
      { id: "Outside", ast: parseFormula("=3"), cell: { sheet: "Outside", row: 1, column: 1 } },
    ];
    const graph = buildDependencyGraph(nodes);
    expect([...graph.edges.get("Summary")!].sort()).toEqual(["Last", "Middle", "Outside"]);
  });

  it("orders deep dependency chains without recursive-stack failure", () => {
    const size = 3_000;
    const nodes = Array.from({ length: size }, (_, index) => ({
      id: `A${index + 1}`,
      ast: parseFormula(index === size - 1 ? "=1" : `=A${index + 2}+1`),
      cell: { sheet: "Sheet1", row: index + 1, column: 1 },
    }));
    const order = topologicalFormulaOrder(buildDependencyGraph(nodes));
    expect(order).toHaveLength(size);
    expect(order[0]).toBe(`A${size}`);
    expect(order.at(-1)).toBe("A1");
  });
});

describe("formula evaluator", () => {
  const workbook: FormulaWorkbookContext = {
    getSheetId: (name) => (name === "Data" ? 1 : name === "Sheet1" ? 0 : undefined),
    getSheetName: (id) => (id === 0 ? "Sheet1" : id === 1 ? "Data" : undefined),
    getCell: (_sheet, row, column) => ({ value: row * 10 + column }),
    getDefinedFormula: () => undefined,
    getTableName: () => undefined,
    getRowCount: () => 100,
    getColumnCount: () => 100,
  };

  it("evaluates arithmetic, functions, and references", () => {
    expect(evaluateFormula("=1+2*3", workbook).value).toBe(7);
    expect(evaluateFormula("=SUM(A1:B2)", workbook).value).toBe(66);
    expect(evaluateFormula('=IF(A1>5,"high","low")', workbook).value).toBe("high");
    expect(evaluateFormula("=Data!A1", workbook).value).toBe(11);
  });

  it("propagates Excel errors for invalid operations", () => {
    expect(evaluateFormula("=1/0", workbook).value).toBe("#DIV/0!");
    expect(evaluateFormula('="1"+1', workbook).value).toBe(2);
    expect(evaluateFormula('=-"x"', workbook).value).toBe("#VALUE!");
  });

  it("compares blank cells with numeric, text, and logical zero values", () => {
    const blankWorkbook: FormulaWorkbookContext = { ...workbook, getCell: () => ({ value: null }) };
    expect(evaluateFormula('=A1=""', blankWorkbook).value).toBe(true);
    expect(evaluateFormula("=A1=0", blankWorkbook).value).toBe(true);
    expect(evaluateFormula('=A1<>"x"', blankWorkbook).value).toBe(true);
  });

  it("separates direct and reference aggregate coercion", () => {
    const mixedWorkbook: FormulaWorkbookContext = {
      ...workbook,
      getCell: (_sheet, row) => ({
        value: row === 1 ? "2" : row === 2 ? true : null,
      }),
    };
    expect(evaluateFormula('=SUM(A1:A2,"3",TRUE)', mixedWorkbook).value).toBe(4);
    expect(evaluateFormula("=AVERAGE(A1:A2)", mixedWorkbook).value).toBe("#DIV/0!");
    expect(evaluateFormula('=SUM("3",TRUE)', mixedWorkbook).value).toBe(4);
    const rangeCoercionWorkbook: FormulaWorkbookContext = {
      ...workbook,
      getCell: (_sheet, row) => ({
        value: row === 1 ? "4" : row === 2 ? true : row === 3 ? "16" : 2,
      }),
    };
    expect(evaluateFormula("=GEOMEAN(A1:A4)", rangeCoercionWorkbook).value).toBe(2);
  });

  it("ignores blank cells in scalar and reference A functions", () => {
    const blankWorkbook: FormulaWorkbookContext = {
      ...workbook,
      getCell: () => ({ value: null }),
    };
    expect(evaluateFormula("=AVERAGEA(A1:A3)", blankWorkbook).value).toBe("#DIV/0!");
    expect(evaluateFormula("=AVERAGEA(A1)", blankWorkbook).value).toBe("#DIV/0!");
    expect(evaluateFormula("=MAXA(A1)", blankWorkbook).value).toBe(0);
    expect(evaluateFormula("=MINA(A1)", blankWorkbook).value).toBe(0);
  });

  it("reduces large extrema without spreading into the call stack", () => {
    const largeWorkbook: FormulaWorkbookContext = {
      ...workbook,
      getRowCount: () => 125_000,
    };
    expect(evaluateFormula("=MAX(A1:A125000)", largeWorkbook).value).toBe(1_250_001);
    expect(evaluateFormula("=MIN(A1:A125000)", largeWorkbook).value).toBe(11);
  });

  it("recognizes runtime errors and defined names consistently", () => {
    expect(evaluateFormula("=IFNA(NA(),1)", workbook).value).toBe(1);
    const namedWorkbook: FormulaWorkbookContext = {
      ...workbook,
      getDefinedFormula: (name) => (name.toUpperCase() === "PI" ? "=3.14" : undefined),
    };
    expect(parseFormula("=PI").kind).toBe("name");
    expect(evaluateFormula("=PI", namedWorkbook).value).toBe(3.14);
  });

  it("counts sparse worksheet ranges without materializing every cell", () => {
    const sparseWorkbook: FormulaWorkbookContext = {
      getSheetId: () => 0,
      getSheetName: () => "Sheet1",
      getCell: () => undefined,
      getDefinedFormula: () => undefined,
      getTableName: () => undefined,
      getRowCount: () => 1_048_576,
      getColumnCount: () => 16_384,
      *getCells(_sheet, bounds) {
        if (bounds.firstRow === 1) yield { row: 1, column: 1, value: "2" };
        if (bounds.lastRow === 1_048_576) yield { row: 1_048_576, column: 16_384, value: true };
      },
    };
    expect(evaluateFormula("=COUNT(A1:XFD1048576)", sparseWorkbook).value).toBe(0);
    expect(evaluateFormula("=COUNTA(A1:XFD1048576)", sparseWorkbook).value).toBe(2);
    const sparseUnionWorkbook: FormulaWorkbookContext = {
      ...sparseWorkbook,
      *getCells(_sheet, bounds) {
        if (bounds.firstRow === 1 && bounds.firstColumn === 3 && bounds.lastColumn >= 3)
          yield { row: 1, column: 3, value: 7 };
      },
    };
    expect(evaluateFormula("=COUNTA((A1:XFD1048576,C1:D2))", sparseUnionWorkbook).value).toBe(1);
  });

  it("uses Excel operator precedence and postfix operators", () => {
    expect(evaluateFormula("=-2^2", workbook).value).toBe(4);
    expect(evaluateFormula("=2^-2", workbook).value).toBe(0.25);
    expect(evaluateFormula("=-2%", workbook).value).toBeCloseTo(-0.02);
    expect(evaluateFormula("=2%^2", workbook).value).toBeCloseTo(0.0004);
  });

  it("propagates errors stored in referenced cells", () => {
    const errorWorkbook: FormulaWorkbookContext = {
      ...workbook,
      getCell: (_sheet, row, column) =>
        row === 1 && column === 1 ? { value: "#N/A", error: true } : { value: 1 },
    };
    expect(evaluateFormula("=SUM(A1:B1)", errorWorkbook).value).toBe("#N/A");
    expect(evaluateFormula("=IFERROR(A1,2)", errorWorkbook).value).toBe(2);
    expect(evaluateFormula("=ISERROR(A1)", errorWorkbook).value).toBe(true);
  });

  it("classifies error values in counting and information functions", () => {
    const errorWorkbook: FormulaWorkbookContext = {
      ...workbook,
      getCell: (_sheet, row, column) =>
        row === 1 && column === 1 ? { value: "#N/A", error: true } : { value: 1 },
    };
    expect(evaluateFormula("=COUNT(A1)", errorWorkbook).value).toBe(0);
    expect(evaluateFormula("=COUNTA(A1)", errorWorkbook).value).toBe(1);
    expect(evaluateFormula("=COUNT(1/0)", workbook).value).toBe(0);
    expect(evaluateFormula("=AND({TRUE,FALSE})", workbook).value).toBe(false);
    expect(evaluateFormula("=OR({FALSE,TRUE})", workbook).value).toBe(true);
    expect(evaluateFormula('=AND({"x",1})', workbook).value).toBe(true);
    expect(evaluateFormula('=AND({"x"})', workbook).value).toBe("#VALUE!");
  });

  it("evaluates whole-column, 3D, named, table, and union ranges", () => {
    const richWorkbook: FormulaWorkbookContext = {
      getSheetId: (name) => (name === "Sheet1" ? 0 : name === "Sheet2" ? 1 : undefined),
      getSheetName: (id) => (id === 0 ? "Sheet1" : id === 1 ? "Sheet2" : undefined),
      getCell: (sheet, row, column) => ({ value: sheet * 100 + row * 10 + column }),
      getDefinedFormula: (name) => (name.toUpperCase() === "RATE" ? "=A1*2" : undefined),
      getTableName: (name) =>
        name === "Sales"
          ? {
              id: 0,
              sheetId: 0,
              range: { firstRow: 1, firstColumn: 1, lastRow: 2, lastColumn: 1 },
              columnNames: ["Alpha", "Beta"],
            }
          : undefined,
      getRowCount: (sheet) => (sheet === 0 ? 2 : 3),
      getColumnCount: () => 3,
    };
    expect(evaluateFormula("=SUM(A:A)", richWorkbook).value).toBe(32);
    expect(evaluateFormula("=SUM(5:7)", richWorkbook).value).toBe(558);
    expect(evaluateFormula("=ROW(5:7)", richWorkbook).value).toEqual([[5], [6], [7]]);
    expect(evaluateFormula("=SUM(Sheet1:Sheet2!A1:A1)", richWorkbook).value).toBe(122);
    expect(evaluateFormula("=Rate", richWorkbook).value).toBe(22);
    expect(evaluateFormula("=rate", richWorkbook).value).toBe(22);
    expect(evaluateFormula("=SUM(Sales)", richWorkbook).value).toBe(32);
    expect(evaluateFormula("=SUM((A1,B1))", richWorkbook).value).toBe(23);
    expect(evaluateFormula("=SUM(Sales[Alpha])", richWorkbook).value).toBe(32);
    const aggregateTableWorkbook: FormulaWorkbookContext = {
      ...richWorkbook,
      isRowHidden: (_sheet, row) => row === 2,
      getTableName: (name) =>
        name === "Sales"
          ? {
              id: 0,
              sheetId: 0,
              range: { firstRow: 1, firstColumn: 1, lastRow: 3, lastColumn: 1 },
              columnNames: ["Alpha", "Beta"],
              hasHeaders: true,
            }
          : undefined,
      getCell: (_sheet, row, column) =>
        row === 1
          ? { value: `Header${column}` }
          : row === 3
            ? { value: 100, formula: "=SUBTOTAL(9,A1)" }
            : { value: 10 },
    };
    expect(evaluateFormula("=AGGREGATE(9,1,Sales[Alpha])", aggregateTableWorkbook).value).toBe(0);
    expect(evaluateFormula("=AGGREGATE(9,5,Sales[Alpha])", aggregateTableWorkbook).value).toBe(100);
    expect(evaluateFormula("=SUM(Sales[alpha])", richWorkbook).value).toBe(32);
    expect(evaluateFormula("=SUM(Sales[[Alpha]:[Beta]])", richWorkbook).value).toBe(66);
    expect(evaluateFormula("=ROWS(Sales[#All])", richWorkbook).value).toBe(2);
  });

  it("intersects defined names and matches table columns case-insensitively", () => {
    const namedWorkbook: FormulaWorkbookContext = {
      ...workbook,
      getDefinedFormula: (name) => (name.toUpperCase() === "MYRANGE" ? "=A1:A3" : undefined),
    };
    expect(evaluateFormula("=@MyRange", namedWorkbook, { row: 2, column: 2 }).value).toBe(21);
  });

  it("evaluates nonadjacent table columns without filling gaps", () => {
    const tableWorkbook: FormulaWorkbookContext = {
      ...workbook,
      getTableName: (name) =>
        name === "Sales"
          ? {
              id: 0,
              sheetId: 0,
              range: { firstRow: 1, firstColumn: 5, lastRow: 2, lastColumn: 7 },
              columnNames: ["Alpha", "Beta", "Gamma"],
            }
          : undefined,
    };
    expect(evaluateFormula("=Sales[[Alpha],[Gamma]]", tableWorkbook).value).toEqual([
      [15, 17],
      [25, 27],
    ]);
  });

  it("limits nonadjacent table column unions", () => {
    const largeTableWorkbook: FormulaWorkbookContext = {
      ...workbook,
      getCell: () => undefined,
      getTableName: (name) =>
        name === "Sales"
          ? {
              id: 0,
              sheetId: 0,
              range: { firstRow: 1, firstColumn: 1, lastRow: 1_000_001, lastColumn: 2 },
              columnNames: ["Alpha", "Beta"],
            }
          : undefined,
      getRowCount: () => 1_000_001,
      getColumnCount: () => 2,
    };
    expect(evaluateFormula("=Sales[[Alpha],[Beta]]", largeTableWorkbook).value).toBe("#VALUE!");
  });

  it("intersects implicit references at the calling row or column", () => {
    expect(evaluateFormula("=@A1:A3", workbook, { row: 2, column: 2 }).value).toBe(21);
    expect(evaluateFormula("=@A1:C1", workbook, { row: 2, column: 2 }).value).toBe(12);
    expect(evaluateFormula("=@A1:C3", workbook, { row: 2, column: 2 }).value).toBe("#VALUE!");
    expect(evaluateFormula("=@A1:A3", workbook).value).toBe("#VALUE!");
  });

  it("evaluates subtotal references without a spreadsheet host", () => {
    const hiddenWorkbook: FormulaWorkbookContext = {
      ...workbook,
      getCell: (sheet, row, column) =>
        sheet === 0 && row === 1 && column === 1
          ? { value: 1, formula: "=SUBTOTAL(9,B1)" }
          : { value: row * 10 + column },
      isRowHidden: (_sheet, row) => row === 2,
    };
    expect(evaluateFormula("=SUBTOTAL(9,A1:A2)", hiddenWorkbook).value).toBe(21);
    expect(evaluateFormula("=SUBTOTAL(9,C1:D1)", hiddenWorkbook).value).toBe(27);
    expect(evaluateFormula("=SUBTOTAL(109,A1:A2)", hiddenWorkbook).value).toBe(0);
    expect(evaluateFormula("=SUBTOTAL(109,A1:B2)", hiddenWorkbook).value).toBe(12);
    expect(evaluateFormula("=SUBTOTAL(3,A1:A2)", hiddenWorkbook).value).toBe(1);
    const denseWorkbook: FormulaWorkbookContext = {
      ...workbook,
      getSheetId: (name) => (name === "S1" ? 0 : name === "S2" ? 1 : undefined),
      getSheetName: (id) => (id === 0 ? "S1" : id === 1 ? "S2" : undefined),
      getCell: () => undefined,
      getRowCount: () => 1_048_576,
      getColumnCount: () => 16_384,
    };
    expect(evaluateFormula("=SUBTOTAL(3,S1:S2!A:A)", denseWorkbook).value).toBe("#VALUE!");
  });

  it("broadcasts compatible matrices in binary operations", () => {
    expect(evaluateFormula("=A1:A2+B1:B2", workbook).value).toEqual([[23], [43]]);
  });

  it("broadcasts scalar functions with modern array semantics", () => {
    expect(evaluateFormula("=ABS({-1,2})", workbook).value).toEqual([[1, 2]]);
    expect(evaluateFormula('=TEXT({1,2},"0")', workbook).value).toEqual([["1", "2"]]);
    expect(evaluateFormula("=ROUND({1.23,2.34},1)", workbook).value).toEqual([[1.2, 2.3]]);
    expect(evaluateFormula("=IF({TRUE,FALSE},1,2)", workbook).value).toEqual([[1, 2]]);
    expect(evaluateFormula("=IF({TRUE,FALSE},{1,2;3,4},0)", workbook).value).toEqual([
      [1, 0],
      [3, 0],
    ]);
    expect(evaluateFormula("=IFERROR({1,1/0},9)", workbook).value).toEqual([[1, 9]]);
    expect(evaluateFormula("=IFERROR({1/0,2;3,4},{9,8;7,6})", workbook).value).toEqual([
      [9, 2],
      [3, 4],
    ]);
    expect(evaluateFormula("=IFERROR({1/0,1/0;1/0,1/0},{9,8;7,6})", workbook).value).toEqual([
      [9, 8],
      [7, 6],
    ]);
  });

  it("broadcasts unary operators and preserves omitted IF-branch semantics", () => {
    expect(evaluateFormula("=-{1,2,3}", workbook).value).toEqual([[-1, -2, -3]]);
    expect(evaluateFormula("=--{1,2,3}", workbook).value).toEqual([[1, 2, 3]]);
    expect(evaluateFormula("={1,2,3}%", workbook).value).toEqual([[0.01, 0.02, 0.03]]);
    expect(evaluateFormula("=IF(TRUE,,5)", workbook).value).toBe(0);
    expect(evaluateFormula("=IF(FALSE,1)", workbook).value).toBe(false);
    expect(evaluateFormula("=IF(FALSE,1,)", workbook).value).toBe(false);
    expect(evaluateFormula("=IF({TRUE,FALSE},,5)", workbook).value).toEqual([[0, 5]]);
    expect(evaluateFormula("=IF({TRUE,FALSE},1)", workbook).value).toEqual([[1, false]]);
  });

  it("implements date systems, complex trigonometry, and regression statistics", () => {
    expect(evaluateFormula('=IMCSC("i")', workbook).value).toMatch(/^-0\.850918128/u);
    expect(evaluateFormula('=TEXT(DATE(2026,1,4),"yyyy-mm-dd")', workbook).value).toBe(
      "2026-01-04",
    );
    expect(
      evaluateFormula('=TEXT(DATE(2026,1,4),"yyyy-mm-dd")', workbook, { dateSystem: "1904" }).value,
    ).toBe("2026-01-04");
    expect(evaluateFormula("=CEILING.MATH(-4.3)", workbook).value).toBe(-4);
    expect(evaluateFormula("=CEILING.PRECISE(-4.3,-1)", workbook).value).toBe(-4);
    expect(evaluateFormula("=FLOOR.PRECISE(-4.3)", workbook).value).toBe(-5);
    expect(evaluateFormula("=NORMSINV(0.975)", workbook).value).toBeCloseTo(1.959963984, 7);
    expect(evaluateFormula("=NORM.DIST({1,2},0,1,TRUE)", workbook).value).toHaveLength(1);
    expect(
      (evaluateFormula("=NORM.DIST({1,2},0,1,TRUE)", workbook).value as number[][]).flat(),
    ).toHaveLength(2);
    expect(evaluateFormula('=ISNUMBER({1,"x"})', workbook).value).toEqual([[true, false]]);
    expect(evaluateFormula("=ISNUMBER(1/0)", workbook).value).toBe(false);
    expect(evaluateFormula("=TYPE(1/0)", workbook).value).toBe(16);
    expect(evaluateFormula("=DAY(60.5)", workbook).value).toBe(29);
    expect(evaluateFormula("=EDATE(DATE(2024,1,15)+0.5,1)", workbook).value).toBe(45337.5);
    expect(evaluateFormula("=ROMAN(499,1)", workbook).value).toBe("LDVLIV");
    expect(evaluateFormula('=CONVERT(1,"stone","lbm")', workbook).value).toBeCloseTo(14, 10);
    expect(evaluateFormula("=FACT(7.9)", workbook).value).toBe(5040);
    expect(evaluateFormula("=FACTDOUBLE(7.9)", workbook).value).toBe(105);
    expect(evaluateFormula("=COMBIN(7.9,2.9)", workbook).value).toBe(21);
    expect(evaluateFormula("=GCD(6.9,4.9)", workbook).value).toBe(2);
    expect(evaluateFormula("=LCM(2.9,4.9)", workbook).value).toBe(4);
    expect(evaluateFormula("=GCD(-1,2)", workbook).value).toBe("#NUM!");
    expect(evaluateFormula("=LCM(-1,2)", workbook).value).toBe("#NUM!");
    expect(
      (evaluateFormula("=COMBIN(200,100)", workbook).value as number) / 9.054851465610328e58,
    ).toBeCloseTo(1, 14);
    expect(evaluateFormula("=COMBINA(3,2)", workbook).value).toBe(6);
    expect(evaluateFormula("=MROUND(-5,2)", workbook).value).toBe(-6);
    expect(evaluateFormula("=MROUND(5,-2)", workbook).value).toBe("#NUM!");
    const statistics = evaluateFormula("=LINEST({2;4;6},{1;2;3},TRUE,TRUE)", workbook)
      .value as number[][];
    expect(statistics[0]?.[0]).toBeCloseTo(2, 14);
    expect(statistics[0]?.[1]).toBeCloseTo(0, 14);
    expect(statistics[2]?.[0]).toBeCloseTo(1, 12);
    expect(statistics[3]?.[1]).toBe(1);
    expect(statistics[3]?.[0]).toBe("#NUM!");
    expect(evaluateFormula("=FORECAST.LINEAR(3,{2,4},{1,2})", workbook).value).toBe(6);
    const pairedLine = evaluateFormula('=LINEST({1,"",3},{1,2,3})', workbook).value as number[][];
    expect(pairedLine[0]?.[0]).toBeCloseTo(1, 14);
    expect(pairedLine[0]?.[1]).toBeCloseTo(0, 14);
    expect(evaluateFormula('=TREND({1,"",3},{1,2,3},2)', workbook).value).toEqual([[2]]);
    expect(evaluateFormula('=SLOPE({1,"",3},{1,2,3})', workbook).value).toBe(1);
    expect(evaluateFormula('=CORREL({1,"",3},{1,2,3})', workbook).value).toBe(1);
    expect(evaluateFormula("=TREND({2,4},{1,2},{3})", workbook).value).toEqual([[6]]);
    const slopeOnly = evaluateFormula("=LINEST({2;4},{1;2})", workbook).value as number[][];
    expect(slopeOnly[0]?.[0]).toBeCloseTo(2, 14);
    expect(slopeOnly[0]?.[1]).toBeCloseTo(0, 14);
    const collinearCoefficients = evaluateFormula("=LINEST({32;64;96},{1,10;2,20;3,30})", workbook)
      .value as number[][];
    expect(collinearCoefficients[0]?.[0]).toBeCloseTo(0, 12);
    expect(collinearCoefficients[0]?.[1]).toBeCloseTo(32, 14);
    expect(collinearCoefficients[0]?.[2]).toBeCloseTo(0, 12);
    const collinearPrediction = evaluateFormula(
      "=TREND({32;64;96},{1,10;2,20;3,30},{4,40})",
      workbook,
    ).value as number[][];
    expect(collinearPrediction[0]?.[0]).toBeCloseTo(128, 14);
    const collinearStatistics = evaluateFormula(
      "=LINEST({32;64;96},{1,10;2,20;3,30},TRUE,TRUE)",
      workbook,
    ).value as number[][];
    expect(collinearStatistics[0]?.[0]).toBeCloseTo(0, 12);
    expect(collinearStatistics[0]?.[1]).toBeCloseTo(32, 12);
    expect(collinearStatistics[0]?.[2]).toBeCloseTo(0, 12);
    expect(collinearStatistics[1]?.every((value) => Math.abs(value) < 1e-12)).toBe(true);
    expect(collinearStatistics[2]?.[0]).toBeCloseTo(1, 14);
    expect(collinearStatistics[2]?.[1]).toBeCloseTo(0, 14);
    expect(collinearStatistics[3]?.[0]).toBe("#NUM!");
    expect(collinearStatistics[3]?.[1]).toBe(1);
    expect(collinearStatistics[4]?.[0]).toBeCloseTo(2048, 14);
    expect(collinearStatistics[4]?.[1]).toBeCloseTo(0, 12);
    const skippedColumnStatistics = evaluateFormula(
      "=LINEST({1;3;2;6},{1,0,2;2,0,3;3,0,5;4,0,7},TRUE,TRUE)",
      workbook,
    ).value as number[][];
    expect(skippedColumnStatistics[0]?.[0]).toBeCloseTo(1 / 3, 14);
    expect(skippedColumnStatistics[0]?.[1]).toBe(0);
    expect(skippedColumnStatistics[0]?.[2]).toBeCloseTo(5 / 6, 14);
    expect(skippedColumnStatistics[0]?.[3]).toBeCloseTo(-0.5, 14);
    expect(skippedColumnStatistics[1]?.[0]).toBeCloseTo(Math.sqrt(125 / 9), 14);
    expect(skippedColumnStatistics[1]?.[1]).toBe(0);
    expect(skippedColumnStatistics[1]?.[2]).toBeCloseTo(Math.sqrt(1475 / 36), 12);
    expect(skippedColumnStatistics[1]?.[3]).toBeCloseTo(2.5, 12);
    const tinyCoefficient = evaluateFormula(
      "=LINEST({2E-155;4E-155},{1E-155;2E-155},FALSE)",
      workbook,
    ).value as number[][];
    expect(tinyCoefficient[0]?.[0]).toBeCloseTo(2, 14);
    const subnormalScaleCoefficient = evaluateFormula(
      "=LINEST({2E-200;4E-200},{1E-200;2E-200},FALSE)",
      workbook,
    ).value as number[][];
    expect(subnormalScaleCoefficient[0]?.[0]).toBeCloseTo(2, 14);
    expect(evaluateFormula("=TREND({32;64;96},{1,10},{4})", workbook).value).toBe("#REF!");
    expect(evaluateFormula("=LINEST({32;64;96},{1,10})", workbook).value).toBe("#REF!");
    const growth = evaluateFormula("=GROWTH({32;64;96},{1,10;2,20;3,30},{4,40})", workbook)
      .value as number[][];
    expect(growth[0]?.[0]).toBeCloseTo(174.44357691188551, 12);
    const imperfectStatistics = evaluateFormula("=LINEST({1;3;2},{1;2;3},TRUE,TRUE)", workbook)
      .value as number[][];
    expect(imperfectStatistics[3]?.[0]).toBeCloseTo(0.33333333333333315, 14);
    expect(evaluateFormula("=TREND({1;2},{1;2},3,FALSE)", workbook).value).toEqual([[3]]);
    expect(evaluateFormula("=TREND({3;5},{1,2;2,3},{4,5},FALSE)", workbook).value).toEqual([[9]]);
    const smallScalePrediction = evaluateFormula(
      "=TREND({1;2;3},{0.0000001;0.0000002;0.0000003},0.0000004,FALSE)",
      workbook,
    ).value as number[][];
    expect(smallScalePrediction[0]?.[0]).toBeCloseTo(4, 14);
    const uncenteredStatistics = evaluateFormula("=LINEST({1;3},{1;2},FALSE,TRUE)", workbook)
      .value as number[][];
    expect(uncenteredStatistics[2]?.[0]).toBeCloseTo(0.98, 14);
    const constantPredictorStatistics = evaluateFormula("=LINEST({2;3},{7;7},TRUE,TRUE)", workbook)
      .value as number[][];
    expect(constantPredictorStatistics[0]?.[0]).toBeCloseTo(0, 14);
    expect(constantPredictorStatistics[0]?.[1]).toBeCloseTo(2.5, 14);
    expect(constantPredictorStatistics[1]?.[1]).toBeCloseTo(0.5, 14);
    expect(constantPredictorStatistics[3]?.[1]).toBe(1);
    const interceptStatistics = evaluateFormula(
      "=LINEST({3;4;5;7;6},{1,2;2,1;3,3;4,2;5,4},TRUE,TRUE)",
      workbook,
    ).value as number[][];
    expect(interceptStatistics[1]?.[2]).toBeCloseTo(0.864098759787715, 14);
    const highOffsetStatistics = evaluateFormula(
      "=LINEST({2;4;6},{100000000;100000001;100000002},TRUE,TRUE)",
      workbook,
    ).value as number[][];
    expect(highOffsetStatistics[0]?.[0]).toBeCloseTo(2, 12);
    expect(highOffsetStatistics[0]?.[1]).toBeCloseTo(-199999998, 6);
    expect(highOffsetStatistics[2]?.[0]).toBeCloseTo(1, 14);
    expect(
      evaluateFormula("=TREND({2;4;6},{100000000;100000001;100000002},100000003,TRUE)", workbook)
        .value,
    ).toEqual([[8]]);
    expect(evaluateFormula("=PERCENTRANK.INC({1,2,3,4},2.5)", workbook).value).toBe(0.5);
    expect(evaluateFormula("=PERCENTRANK.EXC({1,2,3,4},2.5)", workbook).value).toBe(0.5);
    expect(evaluateFormula("=PERCENTRANK.EXC({1,2,3,4},2)", workbook).value).toBe(0.4);
    expect(evaluateFormula("=PERCENTRANK.INC({1,2,3,4},2)", workbook).value).toBe(0.333);
    expect(
      stringifyFormula(
        resolveSharedFormula(
          { formula: "A1+B1", masterRow: 1, masterColumn: 1 },
          { row: 2, column: 3 },
        ),
      ),
    ).toBe("=C2+D2");
  });

  it("evaluates reference metadata, compatibility aliases, and text formats", () => {
    expect(evaluateFormula("=AREAS((A1,B1))", workbook).value).toBe(2);
    expect(evaluateFormula("=SUM((A1:B1,C1:C2))", workbook).value).toBe(59);
    expect(evaluateFormula("=SHEETS()", workbook).value).toBe(2);
    expect(evaluateFormula('=CONVERT_OOO(1,"m","ft")', workbook).value).toBeCloseTo(3.280839895, 9);
    expect(evaluateFormula("=EFFECT_ADD(0.12,12)", workbook).value).toBeCloseTo(0.126825, 6);
    expect(evaluateFormula("=ISEVEN_ADD(3)", workbook).value).toBe(false);
    expect(evaluateFormula("=NUMBERSTRING(12,1)", workbook).value).toBe("十二");
    expect(evaluateFormula("=NUMBERSTRING(12,2)", workbook).value).toBe("壹拾贰");
    expect(evaluateFormula('=TEXT(1234.5,"[>1000][Red]0.00")', workbook).value).toBe("1234.50");
    expect(evaluateFormula('=TEXT(1,"#.##")', workbook).value).toBe("1.");
    expect(evaluateFormula('=TEXT(-1234.5,"#,##0;(#,##0)")', workbook).value).toBe("(1,235)");
    expect(evaluateFormula('=TEXT(1.25,"[h]:mm")', workbook).value).toBe("30:01");
  });

  it("filters hidden rows in aggregate references", () => {
    const hiddenWorkbook: FormulaWorkbookContext = {
      ...workbook,
      isRowHidden: (_sheet, row) => row === 2,
    };
    expect(evaluateFormula("=AGGREGATE(9,5,A1:A2)", hiddenWorkbook).value).toBe(11);
    expect(evaluateFormula("=AGGREGATE(9,4,A1:A2)", hiddenWorkbook).value).toBe(32);
    expect(evaluateFormula("=AGGREGATE(9,5,(A1:A2,B1:B2))", hiddenWorkbook).value).toBe(23);
    expect(evaluateFormula("=AGGREGATE(9,5,Sheet1:Data!A1:A1)", hiddenWorkbook).value).toBe(22);
    expect(evaluateFormula("=AGGREGATE(9,4,SUBTOTAL(9,A1:A2))", hiddenWorkbook).value).toBe(32);
    const errorWorkbook: FormulaWorkbookContext = {
      ...workbook,
      getCell: (_sheet, row, column) =>
        row === 1 && column === 1
          ? { value: "#DIV/0!", error: true }
          : { value: row * 10 + column },
    };
    expect(evaluateFormula("=AGGREGATE(9,2,A1:B1)", errorWorkbook).value).toBe(12);
    expect(evaluateFormula("=AGGREGATE(9,4,A1:B1)", errorWorkbook).value).toBe("#DIV/0!");
    const denseWorkbook: FormulaWorkbookContext = {
      ...workbook,
      getSheetId: (name) => (name === "S1" ? 0 : name === "S2" ? 1 : undefined),
      getSheetName: (id) => (id === 0 ? "S1" : id === 1 ? "S2" : undefined),
      getCell: () => undefined,
      getRowCount: () => 1_048_576,
      getColumnCount: () => 16_384,
    };
    expect(evaluateFormula("=AGGREGATE(3,6,S1:S2!A:A)", denseWorkbook).value).toBe("#VALUE!");
  });

  it("evaluates LET and LAMBDA scopes", () => {
    expect(evaluateFormula("=LET(value,6,doubled,value*2,doubled+1)", workbook).value).toBe(13);
    expect(evaluateFormula("=LET(double,LAMBDA(value,value*2),double(6))", workbook).value).toBe(
      12,
    );
    expect(
      evaluateFormula(
        "=LET(fact,LAMBDA(value,IF(value=0,1,value*fact(value-1))),fact(3))",
        workbook,
      ).value,
    ).toBe(6);
    expect(evaluateFormula("=LAMBDA(value,value)", workbook).value).toBe("#CALC!");
    expect(
      evaluateFormula("=MAKEARRAY(2,2,LAMBDA(row,column,row*10+column))", workbook).value,
    ).toEqual([
      [11, 12],
      [21, 22],
    ]);
  });

  it("evaluates LAMBDA iteration helpers", () => {
    expect(evaluateFormula("=BYROW(A1:B2,LAMBDA(row,SUM(row)))", workbook).value).toEqual([
      [23],
      [43],
    ]);
    expect(evaluateFormula("=BYCOL(A1:B2,LAMBDA(column,SUM(column)))", workbook).value).toEqual([
      [32, 34],
    ]);
    expect(evaluateFormula("=MAP(A1:B2,LAMBDA(value,value*2))", workbook).value).toEqual([
      [22, 24],
      [42, 44],
    ]);
    expect(
      evaluateFormula("=REDUCE(0,A1:A2,LAMBDA(total,value,total+value))", workbook).value,
    ).toBe(32);
    expect(
      evaluateFormula("=SCAN(0,A1:A2,LAMBDA(total,value,total+value))", workbook).value,
    ).toEqual([[11], [32]]);
  });

  it("groups and pivots multiple fields", () => {
    expect(evaluateFormula('=GROUPBY({"A";"A";"B"},{1;2;3},SUM)', workbook).value).toEqual([
      ["A", 3],
      ["B", 3],
    ]);
    expect(evaluateFormula('=GROUPBY({"A";"A";"B"},{1;2;3},SUM,0,0,-1)', workbook).value).toEqual([
      ["B", 3],
      ["A", 3],
    ]);
    expect(
      evaluateFormula('=PIVOTBY({"A";"B";"A"},{"X";"Y";"X"},{1;2;3},SUM)', workbook).value,
    ).toEqual([
      ["A", 4, 0],
      ["B", 0, 2],
    ]);
  });

  it("evaluates conditional branches lazily", () => {
    expect(evaluateFormula("=CHOOSE(1,2,1/0)", workbook).value).toBe(2);
    expect(evaluateFormula("=AND(FALSE,1/0)", workbook).value).toBe(false);
    expect(evaluateFormula("=OR(TRUE,1/0)", workbook).value).toBe(true);
    expect(evaluateFormula("=IFS(FALSE,1,TRUE,2)", workbook).value).toBe(2);
    expect(evaluateFormula('=SWITCH(2,1,"one",2,"two","other")', workbook).value).toBe("two");
  });

  it("implements aggregate and weekday variants", () => {
    expect(evaluateFormula("=AGGREGATE(9,6,A1:B2)", workbook).value).toBe(66);
    expect(evaluateFormula("=AGGREGATE(1,6,A1:B2)", workbook).value).toBe(16.5);
    expect(evaluateFormula("=AGGREGATE(12,6,A1:B2)", workbook).value).toBe(16.5);
    expect(evaluateFormula("=AGGREGATE(14,6,A1:B2,2)", workbook).value).toBe(21);
    expect(evaluateFormula("=WEEKDAY(DATE(2026,1,4),2)", workbook).value).toBe(7);
    expect(evaluateFormula("=WEEKDAY(DATE(2026,1,5),12)", workbook).value).toBe(7);
    expect(evaluateFormula("=WEEKDAY(DATE(2026,1,5),99)", workbook).value).toBe("#NUM!");
  });

  it("implements canonical signature-level functions", () => {
    expect(evaluateFormula("=EDATE(DATE(2024,1,15),3)=DATE(2024,4,15)", workbook).value).toBe(true);
    expect(evaluateFormula("=SORTBY({3;1;2},{1;2;3})", workbook).value).toEqual([[3], [1], [2]]);
    expect(evaluateFormula("=SORTBY({3,1,2},{1,2,3},-1)", workbook).value).toEqual([[2, 1, 3]]);
    expect(evaluateFormula('=ADDRESS(1,1,2,TRUE,"Sheet1")', workbook).value).toBe("Sheet1!A$1");
    expect(evaluateFormula("=ADDRESS(1,1,2,FALSE)", workbook).value).toBe("R1C[1]");
    expect(evaluateFormula("=SERIESSUM(2,0,1,{1,2,3})", workbook).value).toBe(17);
    expect(
      evaluateFormula('=GROUPBY({"A";"B";"A"},{1;2;3},LAMBDA(values,SUM(values)),2,1)', workbook)
        .value,
    ).toEqual([
      ["Group", "Result"],
      ["A", 4],
      ["B", 2],
      ["Total", 6],
    ]);
    expect(evaluateFormula('=GROUPBY({"A";"B";"A"},{1;2;3},"SUM")', workbook).value).toEqual([
      ["A", 4],
      ["B", 2],
    ]);
    expect(
      evaluateFormula(
        '=PIVOTBY({"A";"A";"B"},{"X";"Y";"X"},{1;2;3},LAMBDA(values,SUM(values)),2,1,0,1)',
        workbook,
      ).value,
    ).toEqual([
      ["Group", "X", "Y", "Total"],
      ["A", 1, 2, 3],
      ["B", 3, 0, 3],
      ["Total", 4, 2, 6],
    ]);
  });

  it("implements Excel date, lookup, and shape semantics", () => {
    expect(evaluateFormula('=DATEDIF(DATE(2024,1,31),DATE(2024,2,29),"M")', workbook).value).toBe(
      1,
    );
    expect(evaluateFormula('=DATEDIF(DATE(2023,1,15),DATE(2024,3,10),"YD")', workbook).value).toBe(
      55,
    );
    expect(evaluateFormula('=DATEDIF(DATE(2024,2,1),DATE(2024,1,1),"D")', workbook).value).toBe(
      "#NUM!",
    );
    expect(evaluateFormula('=VLOOKUP(1,{{"1",10}},2,FALSE)', workbook).value).toBe("#N/A");
    expect(evaluateFormula("=SUMPRODUCT({1,2},{3;4})", workbook).value).toBe("#VALUE!");
    expect(evaluateFormula('=TEXT(0.5,"0%")', workbook).value).toBe("50%");
    expect(evaluateFormula('=TEXT(1234.5,"#,##0.00")', workbook).value).toBe("1,234.50");
    expect(evaluateFormula('=ARRAYTOTEXT({1,"a"},1)', workbook).value).toBe('{1,"a"}');
    expect(evaluateFormula('=REGEXTEST("Office","office")', workbook).value).toBe(false);
    expect(evaluateFormula('=REGEXTEST("Office","office",1)', workbook).value).toBe(true);
    expect(evaluateFormula('=TEXTAFTER("a-b-c","-",-1)', workbook).value).toBe("c");
    expect(evaluateFormula('=TEXTAFTER("a-b-c","Z",1,0,0,"missing")', workbook).value).toBe(
      "missing",
    );
    expect(evaluateFormula('=TEXTSPLIT("a,b;c,d",",",";")', workbook).value).toEqual([
      ["a", "b"],
      ["c", "d"],
    ]);
    expect(evaluateFormula('=TEXTSPLIT("a,b;c",",",";")', workbook).value).toEqual([
      ["a", "b"],
      ["c", "#N/A"],
    ]);
    expect(evaluateFormula('=TEXTSPLIT("a,b;c",",",";",,,{1;2})', workbook).value).toEqual([
      ["a", "b"],
      ["c", "#N/A"],
    ]);
    expect(evaluateFormula("=WRAPCOLS({1;2;3;4;5;6},2)", workbook).value).toEqual([
      [1, 3, 5],
      [2, 4, 6],
    ]);
    expect(evaluateFormula('=WORKDAY.INTL(DATE(2026,1,1),1,"1111111")', workbook).value).toBe(
      "#NUM!",
    );
    expect(evaluateFormula("=WORKDAY(1,9E+99)", workbook).value).toBe("#NUM!");
    expect(evaluateFormula("=WORKDAY(1,2200000)", workbook).value).toBe("#NUM!");
    expect(evaluateFormula("=TIME(-1,0,0)", workbook).value).toBe("#NUM!");
    expect(evaluateFormula("=TIME(32767,0,0)", workbook).value).toBe(0.2916666666666667);
    expect(evaluateFormula("=TIME(32768,0,0)", workbook).value).toBe("#NUM!");
    expect(evaluateFormula("=TIME(1.9,59.9,59.9)", workbook).value).toBe(0.08332175925925926);
    expect(evaluateFormula('=NETWORKDAYS.INTL(1,7,"2")', workbook).value).toBe(6);
    expect(evaluateFormula('=NETWORKDAYS.INTL(1,8,"7")', workbook).value).toBe(7);
    expect(evaluateFormula("=DAY(60)", workbook).value).toBe(29);
    expect(evaluateFormula('=DATEVALUE("2026-10-26")', workbook).value).toBe(46321);
    expect(evaluateFormula('=DATEVALUE("2026/10/26")', workbook).value).toBe(46321);
    expect(evaluateFormula('=TEXT(60,"yyyy-mm-dd")', workbook).value).toBe("1900-02-29");
    expect(evaluateFormula('=TEXT(60,"dddd")', workbook).value).toBe("Wednesday");
    expect(evaluateFormula('=TEXT(60,"ddd")', workbook).value).toBe("Wed");
    expect(evaluateFormula("=DATE(100,1,1)", workbook).value).toBe(36526);
    expect(evaluateFormula("=DATE(100,1,1)", { ...workbook, dateSystem: "1904" }).value).toBe(
      35064,
    );
    expect(evaluateFormula("=DATE(1899,12,31)=DATE(3799,12,31)", workbook).value).toBe(true);
    expect(evaluateFormula("=EDATE(60,1)", workbook).value).toBe(89);
    expect(evaluateFormula("=EOMONTH(60,0)", workbook).value).toBe(60);
    expect(evaluateFormula("=EDATE(DATE(1900,2,29),0)", workbook).value).toBe(60);
    expect(evaluateFormula("=EDATE(DATE(1900,2,29),-1)", workbook).value).toBe(31);
    expect(evaluateFormula("=EOMONTH(DATE(1900,2,29),-1)", workbook).value).toBe(31);
    expect(evaluateFormula("=EOMONTH(59,1)", workbook).value).toBe(91);
    expect(evaluateFormula("=EOMONTH(DATE(1900,1,1),1)", workbook).value).toBe(60);
    expect(evaluateFormula("=-{1,2,3}", workbook).value).toEqual([[-1, -2, -3]]);
    expect(evaluateFormula("=--{1,2,3}", workbook).value).toEqual([[1, 2, 3]]);
    expect(evaluateFormula("={1,2,3}%", workbook).value).toEqual([[0.01, 0.02, 0.03]]);
    expect(evaluateFormula("=INDEX({1,2,3;4,5,6},2)", workbook).value).toEqual([[4, 5, 6]]);
    expect(evaluateFormula("=INDEX({1,2,3},2)", workbook).value).toBe(2);
    expect(evaluateFormula("=INDEX(A1:A3,2)", workbook).value).toBe(21);
    expect(evaluateFormula("=INDEX(A1:B2,0,0)", workbook).value).toEqual([
      [11, 12],
      [21, 22],
    ]);
    expect(evaluateFormula('=NUMBERVALUE("abc")', workbook).value).toBe("#VALUE!");
    expect(evaluateFormula('=NUMBERVALUE("2,50",",")', workbook).value).toBe(2.5);
    expect(evaluateFormula('=REGEXREPLACE("a b a b","a","X",0)', workbook).value).toBe("X b X b");
    expect(evaluateFormula('=REGEXREPLACE("ab","(a)(b)","$2$1",1)', workbook).value).toBe("ba");
    expect(evaluateFormula('=AVERAGEIF({"a";"b"},"*",{"x";"y"})', workbook).value).toBe("#DIV/0!");
    expect(evaluateFormula("=YEARFRAC(DATE(2024,1,31),DATE(2024,3,15),0)", workbook).value).toBe(
      0.125,
    );
    expect(evaluateFormula("=YEARFRAC(DATE(2024,3,15),DATE(2024,1,31),0)", workbook).value).toBe(
      "#NUM!",
    );
    expect(evaluateFormula("=DATE(1900,2,29)", workbook).value).toBe(60);
    expect(evaluateFormula('=WORKDAY.INTL(DATE(2026,1,1),1,"99")', workbook).value).toBe("#NUM!");
    expect(evaluateFormula("=INDEX({1,2;3,4},0,2)", workbook).value).toEqual([[2], [4]]);
    expect(evaluateFormula("=INDEX({1,2;3,4},2,0)", workbook).value).toEqual([[3, 4]]);
    expect(evaluateFormula('=TEXTAFTER("A-B","a")', workbook).value).toBe("#N/A");
    expect(evaluateFormula('=TEXTAFTER("A-B","A",1,1)', workbook).value).toBe("-B");
    expect(evaluateFormula('=TEXTBEFORE("İa","ia",1,1)', workbook).value).toBe("");
    expect(evaluateFormula('=TEXTBEFORE("aİb","b",1,1)', workbook).value).toBe("aİ");
    expect(evaluateFormula('=SEARCH("~*","a*b")', workbook).value).toBe(2);
    expect(evaluateFormula('=FINDB("b","中文b",5)', workbook).value).toBe(5);
    expect(evaluateFormula('=TEXT(TIME(0,5,9),"mm:ss")', workbook).value).toBe("05:09");
    expect(evaluateFormula("=EXPAND({1,2},1)", workbook).value).toEqual([[1, 2]]);
    expect(evaluateFormula("=EXPAND({1,2},1,1)", workbook).value).toBe("#VALUE!");
    expect(evaluateFormula("=POISSON(2,3,FALSE)", workbook).value).toBeCloseTo(
      0.224041807655388,
      12,
    );
  });

  it("serializes Excel 1900 and 1904 date systems", () => {
    expect(evaluateFormula("=DATE(1900,1,1)", workbook).value).toBe(1);
    expect(evaluateFormula("=DATE(1900,2,28)", workbook).value).toBe(59);
    expect(evaluateFormula("=DATE(1900,3,1)", workbook).value).toBe(61);
    expect(evaluateFormula("=DATE(2026,1,4)", workbook).value).toBe(46026);
    expect(evaluateFormula("=DATE(2026,1,4)", { ...workbook, dateSystem: "1904" }).value).toBe(
      44564,
    );
    const system1904 = { ...workbook, dateSystem: "1904" } as const;
    expect(evaluateFormula("=DAY(DATE(2026,1,4))", system1904).value).toBe(4);
    expect(evaluateFormula("=MONTH(DATE(2026,1,4))", system1904).value).toBe(1);
    expect(evaluateFormula("=YEAR(DATE(2026,1,4))", system1904).value).toBe(2026);
    expect(evaluateFormula("=ISOWEEKNUM(DATE(2026,1,4))", workbook).value).toBe(1);
  });

  it("applies criteria and pivot selectors", () => {
    expect(evaluateFormula('=COUNTIF(A1:B2,">15")', workbook).value).toBe(2);
    expect(evaluateFormula('=SUMIF(A1:B2,">15")', workbook).value).toBe(43);
    expect(evaluateFormula('=COUNTIF(A1:B2,"1?")', workbook).value).toBe(2);
    expect(evaluateFormula('=COUNTIF(A1:B2,"<>11")', workbook).value).toBe(3);
    const pivotWorkbook: FormulaWorkbookContext = {
      ...workbook,
      getPivotValue: (request) =>
        request.selectors.map(({ field, item }) => `${field}:${item}`).join("|"),
    };
    expect(
      evaluateFormula(
        '=GETPIVOTDATA("Amount","Pivot","Region","East","Product","Alpha")',
        pivotWorkbook,
      ).value,
    ).toBe("Region:East|Product:Alpha");
  });

  it("evaluates reference position and structure functions", () => {
    expect(evaluateFormula("=ROW()", workbook, { row: 7 }).value).toBe(7);
    expect(evaluateFormula("=COLUMN()", workbook, { column: 5 }).value).toBe(5);
    expect(evaluateFormula("=ROW(A2:B3)", workbook).value).toEqual([[2], [3]]);
    expect(evaluateFormula("=COLUMN(B2:C3)", workbook).value).toEqual([[2, 3]]);
    expect(evaluateFormula("=ISREF(A1:B2)", workbook).value).toBe(true);
    expect(evaluateFormula("=ISREF(1)", workbook).value).toBe(false);
    expect(evaluateFormula("=ROW(A2:B3,1)", workbook).value).toBe("#VALUE!");
    expect(evaluateFormula("=OFFSET(A1,1)", workbook).value).toBe("#VALUE!");
    expect(evaluateFormula("=ISREF()", workbook).value).toBe("#VALUE!");
  });

  it("evaluates reference intersection and OpenFormula union", () => {
    expect(evaluateFormula("=A1:B2 A2:B3", workbook).value).toEqual([[21, 22]]);
    const unionAst = parseFormula("=[.A1]~[.B2]", { dialect: "openformula" });
    expect(evaluateFormulaAst(unionAst, workbook)).toEqual([[11], [22]]);
  });

  it("evaluates indirect and offset references", () => {
    expect(evaluateFormula('=INDIRECT("A2")', workbook).value).toBe(21);
    expect(evaluateFormula('=INDIRECT("Data!A1")', workbook).value).toBe(11);
    expect(evaluateFormula('=INDIRECT("R2C2",FALSE)', workbook).value).toBe(22);
    expect(
      evaluateFormula('=INDIRECT("R[0]C[1]",FALSE)', workbook, { row: 2, column: 1 }).value,
    ).toBe(22);
    expect(evaluateFormula("=OFFSET(A1,1,1)", workbook).value).toBe(22);
    expect(evaluateFormula("=OFFSET(A1,1,1,2,1)", workbook).value).toEqual([[22], [32]]);
    expect(evaluateFormula('=INDIRECT("not a reference")', workbook).value).toBe("#REF!");
    expect(evaluateFormula("=OFFSET(A1,100,0)", workbook).value).toBe("#REF!");
  });

  it("evaluates spill and external workbook references", () => {
    const spillWorkbook: FormulaWorkbookContext = {
      ...workbook,
      getSpillRange: () => ({ firstRow: 1, firstColumn: 1, lastRow: 2, lastColumn: 2 }),
    };
    expect(evaluateFormula("=A1#", spillWorkbook).value).toEqual([
      [11, 12],
      [21, 22],
    ]);
    const external = evaluateFormula("=[Book1]Sheet1!A1", {
      ...workbook,
      getWorkbook: (source) =>
        (typeof source === "string" ? source : source.name) === "Book1"
          ? { ...workbook, getCell: (_sheet, row, column) => ({ value: row * 100 + column }) }
          : undefined,
    });
    expect(external.value).toBe(101);
    expect(
      evaluateFormula("=[Missing]Sheet1!A1", { ...workbook, getWorkbook: () => undefined }).value,
    ).toBe("#REF!");
  });

  it("collects dynamic and workbook-qualified dependencies", () => {
    expect(
      collectDependencies(parseFormula("=A1#")).map((dependency) =>
        formatFormulaDependencyTarget(dependency.target),
      ),
    ).toEqual(["A1#"]);
    expect(
      collectDependencies(parseFormula("=[Book1]Sheet1!A1")).map((dependency) =>
        formatFormulaDependencyTarget(dependency.target),
      ),
    ).toEqual(["[Book1]Sheet1!A1"]);
    expect(
      collectDependencies(parseFormula('=INDIRECT("A2")')).map((dependency) =>
        formatFormulaDependencyTarget(dependency.target),
      ),
    ).toEqual(["A2"]);
    expect(
      collectDependencies(parseFormula("=OFFSET(A1,1,0)")).map((dependency) =>
        formatFormulaDependencyTarget(dependency.target),
      ),
    ).toEqual(["A2"]);
    expect(
      collectDependencies(parseFormula("=LET(value,A1,value + B1)")).map((dependency) =>
        formatFormulaDependencyTarget(dependency.target),
      ),
    ).toEqual(["A1", "B1"]);
    expect(
      collectDependencies(parseFormula("=LAMBDA(value,value + 1)(1)")).map((dependency) =>
        formatFormulaDependencyTarget(dependency.target),
      ),
    ).toEqual([]);
    expect(
      collectDependencies(parseFormula('=SUM(INDIRECT("A2"))')).map((dependency) =>
        formatFormulaDependencyTarget(dependency.target),
      ),
    ).toEqual(["A2"]);
    expect(
      collectDependencies(parseFormula('=LET(value,INDIRECT("A2"),value)')).map((dependency) =>
        formatFormulaDependencyTarget(dependency.target),
      ),
    ).toEqual(["A2"]);
    expect(collectDependencies(parseFormula("=OFFSET(A1,1,1,2,2)")).at(-1)?.target).toMatchObject({
      from: { row: 2, column: 2 },
      to: { row: 3, column: 3 },
    });
  });

  it("rewrites dialects, spill references, and storage-level function names", () => {
    const openFormula = parseFormula("=[Sheet1.A1:Sheet2.B2]", { dialect: "openformula" });
    expect(shiftFormulaAst(openFormula, { rows: 1, columns: 1 }).raw).toBe(
      "=[Sheet1.B2:Sheet2.C3]",
    );
    expect(shiftFormulaAst(parseFormula("=A1#"), { rows: 1 }).raw).toBe("=A2#");
    expect(shiftFormulaAst(parseFormula("='Sheet1'!A1#"), { rows: 1 }).raw).toBe("=Sheet1!A2#");
    expect(parseFormula("='Sheet1'!A1#")).toMatchObject({ reference: { spill: true } });
    expect(getFunction("_xlfn._xlws.FILTER")).toBeDefined();
    expect(getFunction("ORG.OPENOFFICE.STDEV")).toBeDefined();
    expect(evaluateFormula("=XNPV(0.1,{-100;50;60},{45292;45474;45657})", workbook).value).toBe(
      evaluateFormula("=XNPV(0.1,{-100;50;60},{46754;46936;47119})", workbook, {
        dateSystem: "1904",
      }).value,
    );
  });

  it("reads formulas through reference snapshots", () => {
    const formulaWorkbook: FormulaWorkbookContext = {
      ...workbook,
      getCell: (_sheet, row, column) =>
        row === 2 && column === 1 ? { value: 2, formula: "=A1+1" } : { value: 1 },
    };
    expect(evaluateFormula("=FORMULATEXT(A2)", formulaWorkbook).value).toBe("=A1+1");
    expect(evaluateFormula("=FORMULA(A2)", formulaWorkbook).value).toBe("=A1+1");
    expect(evaluateFormula("=FORMULATEXT(A1)", formulaWorkbook).value).toBe("#N/A");
  });

  it("delegates host-dependent dynamic and lookup functions", () => {
    const dynamicWorkbook: FormulaWorkbookContext = {
      ...workbook,
      resolveDynamic: (request) => (request.name === "ANCHORARRAY" ? 42 : undefined),
    };
    const lookupWorkbook: FormulaWorkbookContext = {
      ...workbook,
      resolveLookup: (request) => (request.name === "CELL" ? 24 : undefined),
    };
    expect(evaluateFormula("=ANCHORARRAY(A1)", dynamicWorkbook).value).toBe(42);
    expect(evaluateFormula('=CELL("ADDRESS")', lookupWorkbook).value).toBe(24);
    expect(evaluateFormula('=CELL("ADDRESS")', workbook).value).toBe("#REF!");
  });
});

describe("formula function catalog", () => {
  const workbook: FormulaWorkbookContext = {
    getSheetId: () => undefined,
    getSheetName: () => undefined,
    getCell: () => ({ value: null }),
    getDefinedFormula: () => undefined,
    getTableName: () => undefined,
    getRowCount: () => 10,
    getColumnCount: () => 10,
  };

  const hostWorkbook: FormulaWorkbookContext = {
    ...workbook,
    getImage: ({ source, altText, sizing, height, width }) =>
      source === "source" && altText === "alt" && sizing === 0 && height === 100 && width === 50
        ? `image:${source}`
        : undefined,
    getStockHistory: ({ symbol, startDate, endDate, interval, headers, properties, identifier }) =>
      symbol === "MSFT" &&
      startDate === 1 &&
      endDate === 2 &&
      interval === 0 &&
      headers === true &&
      properties?.join(",") === "0,1" &&
      identifier === "id"
        ? [
            ["date", "close"],
            [1, 2],
          ]
        : undefined,
    translateText: ({ sourceLanguage, targetLanguage, text }) =>
      `${String(sourceLanguage)}:${String(targetLanguage)}:${String(text)}`,
    detectLanguage: ({ text }) => (text[0] === "hello" ? "en" : "fr"),
    getFieldValue: ({ value, field }) =>
      value === "record" && field === "name" ? "Ada" : undefined,
    registerMacro: ({ library, procedure }) => `${library}:${procedure}`,
  };

  it("catalogs every researched canonical function", () => {
    expect(formulaFunctions.length).toBe(667);
    expect(listFunctionNames().length).toBe(667);
    for (const entry of formulaFunctions) {
      const metadata = getFunctionMetadata(entry.name);
      expect(JSON.parse(JSON.stringify(metadata))).toEqual(metadata);
      expect(metadata).toMatchObject({
        minArgs: entry.minArgs,
        maxArgs: Number.isFinite(entry.maxArgs) ? entry.maxArgs : "variadic",
      });
      expect(Boolean(metadata?.volatile)).toBe(Boolean(entry.volatile));
      expect(Boolean(metadata?.dynamicDependencies)).toBe(Boolean(entry.dynamicDependencies));
      expect(metadata?.externalContext).toBe(entry.externalContext);
      expect(entry.minArgs).toBeLessThanOrEqual(entry.maxArgs);
    }
  });

  it("registers every canonical function as an executable implementation", () => {
    const catalogNames = formulaFunctions.map((entry) => entry.name);
    const registeredNames = listFunctionNames();
    expect(new Set(registeredNames)).toEqual(new Set(catalogNames));
    expect(registeredNames).toHaveLength(catalogNames.length);
    const context = createFormulaContext(workbook);
    for (const entry of formulaFunctions) {
      const implementation = getFunction(entry.name);
      expect(implementation).toBeDefined();
      let value: unknown;
      try {
        value = implementation?.(
          Array.from({ length: entry.minArgs }, () => null),
          context,
        );
      } catch (error) {
        expect(error, entry.name).toBeInstanceOf(FormulaValueError);
        continue;
      }
      expect(
        value === null ||
          typeof value === "boolean" ||
          typeof value === "number" ||
          typeof value === "string" ||
          value instanceof FormulaError ||
          (Array.isArray(value) &&
            value.every(
              (row) =>
                Array.isArray(row) &&
                row.every(
                  (cell) =>
                    cell === null ||
                    typeof cell === "boolean" ||
                    typeof cell === "number" ||
                    typeof cell === "string" ||
                    cell instanceof FormulaError,
                ),
            )),
      ).toBe(true);
    }
  });

  it("mirrors compatibility aliases and external functions onto real signatures", () => {
    const signaturePairs: readonly [string, string][] = [
      ["PERCENTILEEXC", "PERCENTILE.EXC"],
      ["PERCENTILEINC", "PERCENTILE.INC"],
      ["NETWORKDAYSINTL", "NETWORKDAYS.INTL"],
      ["WORKDAYINTL", "WORKDAY.INTL"],
      ["WEEKNUMOOO", "WEEKNUM"],
      ["RANKAVG", "RANK.AVG"],
      ["RANKEQ", "RANK.EQ"],
      ["WEIBULLDIST", "WEIBULL.DIST"],
      ["ZTESTCOMPAT", "Z.TEST"],
    ];
    for (const [aliasName, targetName] of signaturePairs) {
      expect(getFunctionMetadata(aliasName)).toMatchObject({
        minArgs: getFunctionMetadata(targetName)?.minArgs,
        maxArgs: getFunctionMetadata(targetName)?.maxArgs,
      });
    }
    expect(getFunctionMetadata("CUBEKPIMEMBER")).toMatchObject({ minArgs: 3, maxArgs: 4 });
    expect(getFunctionMetadata("CUBESETCOUNT")).toMatchObject({ minArgs: 1, maxArgs: 1 });
    expect(getFunctionMetadata("CHOOSE")).toMatchObject({ minArgs: 2, maxArgs: 255 });
    expect(getFunctionMetadata("STRINGS")).toMatchObject({
      minArgs: getFunctionMetadata("CONCAT")?.minArgs,
      maxArgs: getFunctionMetadata("CONCAT")?.maxArgs,
    });
    expect(getFunctionMetadata("RAWSUBTRACT")).toMatchObject({ minArgs: 2, maxArgs: 2 });
    expect(getFunctionMetadata("GROUPBY")).toMatchObject({ minArgs: 3, maxArgs: 7 });
    expect(getFunctionMetadata("PIVOTBY")).toMatchObject({ minArgs: 4, maxArgs: 10 });
    expect(getFunctionMetadata("IMAGE")).toMatchObject({ minArgs: 1, maxArgs: 5 });
    expect(getFunctionMetadata("STOCKHISTORY")).toMatchObject({ minArgs: 2, maxArgs: 7 });
    expect(getFunctionMetadata("TRANSLATE")).toMatchObject({ minArgs: 1, maxArgs: 3 });
    expect(getFunctionMetadata("DETECTLANGUAGE")).toMatchObject({ minArgs: 1, maxArgs: 1 });
    expect(getFunctionMetadata("FIELDVALUE")).toMatchObject({ minArgs: 2, maxArgs: 2 });
    expect(getFunctionMetadata("DATESTRING")).toMatchObject({ minArgs: 1, maxArgs: 1 });
    expect(evaluateFormula('=FIELDVALUE("record","name")', hostWorkbook).value).toBe("Ada");
    expect(evaluateFormula('=FIELDVALUE("record","name")', workbook).value).toBe("#REF!");
    expect(evaluateFormula("=DATESTRING(45658)", workbook).value).toBe("2025-01-01");
    expect(evaluateFormula('=DATESTRING("2025-01-01")', workbook).value).toBe("2025-01-01");
    expect(evaluateFormula('=REGISTER.ID("lib","proc")', hostWorkbook).value).toBe("lib:proc");
    expect(getFunctionMetadata("LEGACY.CHIDIST")).toMatchObject({ minArgs: 2, maxArgs: 2 });
    expect(getFunctionMetadata("COUPDAYNC")).toMatchObject({ minArgs: 3, maxArgs: 4 });
    expect(getFunctionMetadata("USDOLLAR")).toMatchObject({ minArgs: 1, maxArgs: 2 });
    expect(evaluateFormula('=TEXTSPLIT("a,b",",")', workbook).value).toEqual([["a", "b"]]);
    expect(evaluateFormula('=VALUE("10000-01-01")', workbook).value).toBe("#NUM!");
    expect(evaluateFormula("=LAMBDA(value,value+1)(1)", workbook).value).toBe(2);
    expect(evaluateFormula('=LAMBDA(value,value & "!")("done")', workbook).value).toBe("done!");
    for (const entry of formulaFunctions) expect(Number.isFinite(entry.maxArgs)).toBe(true);
    const chooseValues = Array.from({ length: 31 }, (_, index) => index + 1).join(",");
    expect(evaluateFormula(`=CHOOSE(31,${chooseValues})`, workbook).value).toBe(31);
    expect(evaluateFormula("=PERCENTILEEXC({1,2,3,4},0.25,99)", workbook).value).toBe("#VALUE!");
    expect(evaluateFormula("=RANKAVG(1,{1,2,3},1,99)", workbook).value).toBe("#VALUE!");
    expect(evaluateFormula("=RAWSUBTRACT(5,2,99)", workbook).value).toBe("#VALUE!");
    expect(evaluateFormula("=WEIBULLDIST(1,1,1,TRUE,99)", workbook).value).toBe("#VALUE!");
    expect(evaluateFormula('=CUBEKPIMEMBER("connection","kpi","value")', workbook).value).toBe(
      "#N/A",
    );
    expect(evaluateFormula('=CUBEKPIMEMBER("connection","kpi")', workbook).value).toBe("#VALUE!");
    expect(evaluateFormula('=CUBESETCOUNT("set")', workbook).value).toBe("#N/A");
    expect(evaluateFormula("=CUBESETCOUNT()", workbook).value).toBe("#VALUE!");
    expect(evaluateFormula('=STRINGS("a","b",{1,2})', workbook).value).toBe("ab12");
    expect(evaluateFormula("=STRINGS()", workbook).value).toBe("#VALUE!");
    expect(evaluateFormula('=CONCAT("a",,"b",{1,2})', workbook).value).toBe("ab12");
    const ifsConditions = Array.from({ length: 127 }, () => "FALSE,1").join(",");
    expect(evaluateFormula(`=IFS(${ifsConditions})`, workbook).value).toBe("#N/A");
    expect(evaluateFormula(`=IFS(${ifsConditions},99)`, workbook).value).toBe("#VALUE!");
    const switchValues = Array.from({ length: 126 }, (_, index) => `${index + 1},"x"`).join(",");
    expect(evaluateFormula(`=SWITCH(255,${switchValues},"fallback")`, workbook).value).toBe(
      "fallback",
    );
    expect(evaluateFormula('=IMAGE("source","alt",0,100,50)', hostWorkbook).value).toBe(
      "image:source",
    );
    expect(
      evaluateFormula('=STOCKHISTORY("MSFT",1,2,0,TRUE,{0,1},"id")', hostWorkbook).value,
    ).toEqual([
      ["date", "close"],
      [1, 2],
    ]);
    expect(evaluateFormula('=TRANSLATE("hello","en","zh")', hostWorkbook).value).toBe(
      "en:zh:hello",
    );
    expect(evaluateFormula('=DETECTLANGUAGE("hello")', hostWorkbook).value).toBe("en");
    expect(evaluateFormula('=TRANSLATE({"a";"b"},"en","zh")', hostWorkbook).value).toEqual([
      ["en:zh:a"],
      ["en:zh:b"],
    ]);
    expect(evaluateFormula('=DETECTLANGUAGE({"hello","world"})', hostWorkbook).value).toEqual([
      ["en", "fr"],
    ]);
    expect(evaluateFormula("=IMAGE(1)", workbook).value).toBe("#REF!");
    expect(evaluateFormula('=STOCKHISTORY("MSFT",1,2,0,TRUE,{6},"id")', hostWorkbook).value).toBe(
      "#VALUE!",
    );
    expect(evaluateFormula('=TRANSLATE(1/0,"en","zh")', workbook).value).toBe("#DIV/0!");
    expect(evaluateFormula("=DETECTLANGUAGE(1/0)", workbook).value).toBe("#DIV/0!");
  });
});

describe("formula errors", () => {
  it("preserves canonical error values without matching error messages", () => {
    const error = new FormulaValueError("#NUM!", "invalid principal", { cause: 0 });
    expect(error.value).toBe("#NUM!");
    expect(error.reason).toBe("invalid principal");
    expect(error.message).toBe("#NUM!: invalid principal");
  });

  it("exposes the complete runtime error model", () => {
    expect(formulaPublic.FormulaError).toBeDefined();
    expect(formulaPublic.isFormulaErrorToken("#N/A")).toBe(true);
    expect(formulaPublic.isFormulaErrorToken("#BUSY!")).toBe(true);
    expect(formulaPublic.isFormulaErrorToken("#NOTE")).toBe(false);
    const error = formulaPublic.normalizeFormulaError("#VALUE!");
    expect(formulaPublic.isFormulaErrorValue(error)).toBe(true);
    expect(formulaPublic.serializeFormulaError(error)).toBe("#VALUE!");
    const implementation = getFunction("ABS")!;
    expect(Object.isFrozen(implementation)).toBe(true);
    expect(Object.isFrozen(implementation.metadata)).toBe(true);
  });

  it("preserves error kinds across nested and broadcast calls", () => {
    const workbook: FormulaWorkbookContext = {
      getSheetId: () => undefined,
      getSheetName: () => undefined,
      getCell: () => undefined,
      getDefinedFormula: () => undefined,
      getTableName: () => undefined,
      getRowCount: () => 1,
      getColumnCount: () => 1,
    };
    expect(evaluateFormula("=ABS(TIME(-1,0,0))", workbook).value).toBe("#NUM!");
    expect(evaluateFormula("=ABS(NOPE())", workbook).value).toBe("#NAME?");
    expect(evaluateFormula("=ABS(ABS(TIME(-1,0,0)))", workbook).value).toBe("#NUM!");
  });

  it("propagates unknown host and numeric-library exceptions", () => {
    const workbook: FormulaWorkbookContext = {
      getSheetId: () => undefined,
      getSheetName: () => undefined,
      getCell: () => undefined,
      getDefinedFormula: () => undefined,
      getTableName: () => undefined,
      getRowCount: () => 1,
      getColumnCount: () => 1,
      getDdeData: () => {
        throw new Error("host unavailable");
      },
    };
    expect(() => evaluateFormula('=DDE("server","file","item")', workbook)).toThrow(
      "host unavailable",
    );
  });
});

describe("formula error regressions", () => {
  const workbook: FormulaWorkbookContext = {
    getSheetId: () => undefined,
    getSheetName: () => undefined,
    getCell: () => ({ value: null }),
    getDefinedFormula: () => undefined,
    getTableName: () => undefined,
    getRowCount: () => 10,
    getColumnCount: () => 10,
  };

  it("treats note-like text as ordinary text", () => {
    expect(evaluateFormula('=N("#NOTE")', workbook).value).toBe(0);
    const noteWorkbook: FormulaWorkbookContext = {
      ...workbook,
      getCell: (_sheet, row, column) =>
        row === 1 && column === 1
          ? { value: "Field" }
          : row === 2 && column === 1
            ? { value: "#NOTE" }
            : row === 1 && column === 3
              ? { value: "Field" }
              : undefined,
    };
    expect(evaluateFormula("=DCOUNTA(A1:A2,1,C1:C2)", noteWorkbook).value).toBe(1);
  });

  it("inspects runtime errors and error literals separately", () => {
    expect(evaluateFormula("=ERROR.TYPE(1/0)", workbook).value).toBe(2);
    const isErr = getFunction("ISERR")!;
    const isNa = getFunction("ISNA")!;
    const context = createFormulaContext(workbook);
    expect(isErr([new FormulaError("#VALUE!")], context)).toBe(true);
    expect(isErr(["#VALUE!"], context)).toBe(false);
    expect(isNa([new FormulaError("#N/A")], context)).toBe(true);
    expect(isNa(["#N/A"], context)).toBe(false);
  });
});

describe("formula review regressions", () => {
  const workbook: FormulaWorkbookContext = {
    getSheetId: (name) => (name === "Sheet1" ? 0 : undefined),
    getSheetName: (id) => (id === 0 ? "Sheet1" : undefined),
    getCell: (_sheet, row, column) => ({ value: row * 10 + column }),
    getDefinedFormula: () => undefined,
    getTableName: () => undefined,
    getRowCount: () => 100,
    getColumnCount: () => 100,
  };

  it("parses unquoted cross-sheet ranges as one 3D reference", () => {
    expect(parseFormula("=Sheet1!A1:Sheet2!B2")).toMatchObject({
      kind: "reference",
      reference: {
        sheets: { names: ["Sheet1", "Sheet2"], range: true },
        areas: [
          {
            fromSheet: "Sheet1",
            toSheet: "Sheet2",
            start: { row: 1, column: 1 },
            end: { row: 2, column: 2 },
          },
        ],
      },
    });
  });

  it("parses whole-column and whole-row 3D references", () => {
    expect(parseFormula("=Sheet1:Sheet2!A:C")).toMatchObject({
      reference: {
        areas: [
          {
            fromSheet: "Sheet1",
            toSheet: "Sheet2",
            start: { kind: "whole-column", column: 1 },
            end: { kind: "whole-column", column: 3 },
          },
        ],
      },
    });
    expect(parseFormula("=Sheet1:Sheet2!1:3")).toMatchObject({
      reference: {
        areas: [
          {
            fromSheet: "Sheet1",
            toSheet: "Sheet2",
            start: { kind: "whole-row", row: 1 },
            end: { kind: "whole-row", row: 3 },
          },
        ],
      },
    });
  });

  it("round-trips plain table selectors and array literals", () => {
    expect(stringifyFormula(parseFormula("=Sales[Amount]"))).toBe("=Sales[Amount]");
    expect(stringifyFormula(parseFormula("=Sales[[#All],[Amount]]"))).toBe(
      "=Sales[[#All],[Amount]]",
    );
    expect(stringifyFormula(parseFormula("={1,2;3,4}"))).toBe("={1,2;3,4}");
    expect(
      stringifyFormula(parseFormula("={1;2|3;4}", { dialect: "openformula" }), {
        dialect: "openformula",
      }),
    ).toBe("={1;2|3;4}");
  });

  it("treats fractional holiday dates as whole days in WORKDAY", () => {
    const holiday = "{DATE(2026,1,5)+0.5}";
    expect(evaluateFormula(`=WORKDAY(DATE(2026,1,2),1,${holiday})`, workbook).value).toBe(46028);
    expect(evaluateFormula(`=WORKDAY(DATE(2026,1,6),-1,${holiday})`, workbook).value).toBe(46024);
  });

  it("normalizes the Excel 1900 pseudo leap day across date inputs", () => {
    expect(evaluateFormula('=DATEVALUE("1900-02-29")', workbook).value).toBe(60);
    expect(evaluateFormula('=WORKDAY(DATE(1900,3,1),-1,{"1900-02-29"})', workbook).value).toBe(59);
  });

  it("ignores blank cells in holiday ranges", () => {
    const holidayWorkbook: FormulaWorkbookContext = {
      ...workbook,
      getCell: (_sheet, row, column) =>
        row === 1 && column === 1 ? { value: 46027 } : { value: null },
    };
    expect(evaluateFormula("=WORKDAY(DATE(2026,1,2),1,A1:A2)", holidayWorkbook).value).toBe(46028);
    expect(
      evaluateFormula("=NETWORKDAYS(DATE(2026,1,2),DATE(2026,1,6),A1:A2)", holidayWorkbook).value,
    ).toBe(2);
  });

  it("passes fifth-review arithmetic and signature regressions", () => {
    expect(parseFormula("=TRUE()").kind).toBe("call");
    expect(parseFormula("=TRUE ()").kind).toBe("call");
    expect(evaluateFormula("=TRUE()", workbook).value).toBe(true);
    expect(evaluateFormula("=TRUE ()", workbook).value).toBe(true);
    expect(evaluateFormula("=FALSE()", workbook).value).toBe(false);
    expect(evaluateFormula('=TEXT(0.125,"# ?/?")', workbook).value).toBe("1/8");
    expect(evaluateFormula('=TEXT(12345,"0.00E+00")', workbook).value).toBe("1.23E+04");
    expect(evaluateFormula("=DEC2BIN(-1)", workbook).value).toBe("1111111111");
    expect(evaluateFormula("=DEC2OCT(-1)", workbook).value).toBe("7777777777");
    expect(evaluateFormula("=DEC2HEX(-1)", workbook).value).toBe("FFFFFFFFFF");
    const nextCoupon = evaluateFormula(
      "=COUPNCD(DATE(2024,1,15),DATE(2024,7,15),2,0)",
      workbook,
    ).value;
    const nextCoupon1904 = evaluateFormula(
      "=COUPNCD(DATE(2024,1,15),DATE(2024,7,15),2,0)",
      workbook,
      { dateSystem: "1904" },
    ).value;
    expect(nextCoupon).toBe(45488);
    expect(nextCoupon1904).toBe(44026);
    expect(evaluateFormula('=LEFT("abc",1,2)', workbook).value).toBe("#VALUE!");
    expect(evaluateFormula('=SUBSTITUTE("abc","a","z",1,2)', workbook).value).toBe("#VALUE!");
    expect(evaluateFormula('=FIND("a","abc",0)', workbook).value).toBe("#VALUE!");
    expect(evaluateFormula('=FIND("a","𐐷a",2)', workbook).value).toBe(2);
    expect(evaluateFormula('=SEARCH("a","𐐷a")', workbook).value).toBe(2);
    expect(evaluateFormula('=SEARCHB("a","𐐷a",2)', workbook).value).toBe("#VALUE!");
    expect(evaluateFormula('=SUBSTITUTE("abc","a","z",0)', workbook).value).toBe("#VALUE!");
    expect(evaluateFormula('=REPT("a",-1)', workbook).value).toBe("#VALUE!");
  });

  it("isolates workbooks, expands spill ranges, and records dynamic offsets", () => {
    const externalGraph = buildDependencyGraph([
      {
        id: "formula",
        ast: parseFormula("=[Book]Sheet1!A1"),
        cell: { sheet: "Sheet1", row: 1, column: 3 },
      },
      { id: "local", ast: parseFormula("=1"), cell: { sheet: "Sheet1", row: 1, column: 1 } },
      {
        id: "external",
        ast: parseFormula("=2"),
        cell: { workbook: "Book", sheet: "Sheet1", row: 1, column: 1 },
      },
    ]);
    expect([...externalGraph.edges.get("formula")!]).toEqual(["external"]);

    const nameGraph = buildDependencyGraph([
      {
        id: "formula",
        ast: parseFormula("=rate"),
        cell: { sheet: "Sheet1", row: 1, column: 4 },
      },
      { id: "RATE", ast: parseFormula("=1"), cell: { sheet: "Sheet1", row: 1, column: 1 } },
    ]);
    expect([...nameGraph.edges.get("formula")!]).toEqual(["RATE"]);

    const spillGraph = buildDependencyGraph(
      [
        {
          id: "formula",
          ast: parseFormula("=SUM(A1#)"),
          cell: { sheet: "Sheet1", row: 1, column: 5 },
        },
        { id: "anchor", ast: parseFormula("=1"), cell: { sheet: "Sheet1", row: 1, column: 2 } },
        { id: "spilled", ast: parseFormula("=2"), cell: { sheet: "Sheet1", row: 2, column: 3 } },
      ],
      {
        getSpillRange: (_workbook, sheet, row, column) =>
          sheet === "Sheet1" && row === 1 && column === 1
            ? { firstRow: 1, firstColumn: 2, lastRow: 2, lastColumn: 3 }
            : undefined,
      },
    );
    expect([...spillGraph.edges.get("formula")!].sort()).toEqual(["anchor", "spilled"]);

    const unknownSpillGraph = buildDependencyGraph([
      {
        id: "formula",
        ast: parseFormula("=SUM(A1#)"),
        cell: { sheet: "Sheet1", row: 1, column: 5 },
      },
      { id: "anchor", ast: parseFormula("=1"), cell: { sheet: "Sheet1", row: 1, column: 2 } },
      { id: "spilled", ast: parseFormula("=2"), cell: { sheet: "Sheet1", row: 2, column: 3 } },
    ]);
    expect([...unknownSpillGraph.edges.get("formula")!].sort()).toEqual(["anchor", "spilled"]);
    expect(
      collectDependencies(parseFormula("=OFFSET(A1,A2,1)")).some(
        ({ target }) => target.kind === "unknown",
      ),
    ).toBe(true);
  });

  it("aggregates multiple AGGREGATE reference arguments", () => {
    expect(evaluateFormula("=AGGREGATE(9,7,A1:B1,C1:D1)", workbook).value).toBe(50);
  });

  it("validates AGGREGATE bounds and named-table implicit intersections", () => {
    expect(evaluateFormula("=AGGREGATE(9,8,A1:A2)", workbook).value).toBe("#VALUE!");
    expect(evaluateFormula("=AGGREGATE(20,0,A1:A2)", workbook).value).toBe("#VALUE!");
    expect(evaluateFormula("=AGGREGATE(14,6,A1:A2,2,3)", workbook).value).toBe("#VALUE!");

    const tableInfo = {
      sheet: "Data",
      range: { firstRow: 1, firstColumn: 5, lastRow: 4, lastColumn: 7 },
      columnNames: ["Alpha", "Beta", "Gamma"],
      hasHeaders: true,
      hasTotals: true,
    };
    const graph = buildDependencyGraph(
      [
        {
          id: "formula",
          ast: parseFormula("=@Sales[Alpha]"),
          cell: { sheet: "Data", row: 3, column: 8 },
        },
        { id: "current", ast: parseFormula("=1"), cell: { sheet: "Data", row: 3, column: 5 } },
        { id: "other", ast: parseFormula("=2"), cell: { sheet: "Data", row: 2, column: 5 } },
      ],
      { getTableName: (name) => (name === "Sales" ? tableInfo : undefined) },
    );
    expect([...graph.edges.get("formula")!]).toEqual(["current"]);
  });

  it("restricts #This Row to the table body", () => {
    const tableWorkbook: FormulaWorkbookContext = {
      ...workbook,
      getTableName: (name) =>
        name === "Sales"
          ? {
              id: 0,
              sheetId: 0,
              range: { firstRow: 2, firstColumn: 1, lastRow: 4, lastColumn: 2 },
              columnNames: ["Alpha", "Beta"],
              hasHeaders: true,
              hasTotals: true,
            }
          : undefined,
    };
    expect(
      evaluateFormula("=Sales[#This Row]", tableWorkbook, { row: 3, column: 1 }).value,
    ).toEqual([[31, 32]]);
    expect(evaluateFormula("=Sales[#This Row]", tableWorkbook, { row: 5, column: 1 }).value).toBe(
      "#VALUE!",
    );
    expect(evaluateFormula("=Sales[#This Row]", tableWorkbook, { row: 2, column: 1 }).value).toBe(
      "#VALUE!",
    );
    expect(evaluateFormula("=Sales[#This Row]", tableWorkbook, { row: 4, column: 1 }).value).toBe(
      "#VALUE!",
    );
    expect(evaluateFormula("=Sales[#Headers]", tableWorkbook).value).toEqual([[21, 22]]);
    expect(evaluateFormula("=Sales[#Totals]", tableWorkbook).value).toEqual([[41, 42]]);
    const emptyBodyWorkbook: FormulaWorkbookContext = {
      ...workbook,
      getTableName: (name) =>
        name === "Sales"
          ? {
              id: 0,
              sheetId: 0,
              range: { firstRow: 1, firstColumn: 1, lastRow: 2, lastColumn: 2 },
              columnNames: ["Alpha", "Beta"],
              hasHeaders: true,
              hasTotals: true,
            }
          : undefined,
    };
    expect(evaluateFormula("=Sales[#Data]", emptyBodyWorkbook).value).toBe("#REF!");
    expect(evaluateFormula("=@Sales[Beta]", tableWorkbook, { row: 3, column: 2 }).value).toBe(32);
  });

  it("distinguishes error snapshots from error-like text", () => {
    const snapshotWorkbook: FormulaWorkbookContext = {
      ...workbook,
      getCell: (_sheet, _row, column) =>
        column === 1 ? { value: "#N/A", error: true } : { value: "#N/A" },
    };
    expect(evaluateFormula("=ISERROR(A1)", snapshotWorkbook).value).toBe(true);
    expect(evaluateFormula("=ISERROR(B1)", snapshotWorkbook).value).toBe(false);
    expect(evaluateFormula("=ISNA(B1)", snapshotWorkbook).value).toBe(false);
    expect(evaluateFormula('=CONCAT("#N/A","x")', workbook).value).toBe("#N/Ax");
    expect(evaluateFormula('=ISERROR("#N/A")', workbook).value).toBe(false);
    expect(evaluateFormula('=IFERROR("#N/A",1)', workbook).value).toBe("#N/A");
  });

  it("broadcasts scalar statistical functions over arrays", () => {
    const result = evaluateFormula("=T.DIST({1;2},2,FALSE)", workbook).value;
    expect(result).toEqual([[expect.any(Number)], [expect.any(Number)]]);
  });

  it("evaluates probability distributions with precise semantics", () => {
    expect(evaluateFormula("=T.DIST(1,2,TRUE)", workbook).value).toBeCloseTo(0.788675134594813, 12);
    expect(evaluateFormula("=T.DIST(1,2,FALSE)", workbook).value).toBeCloseTo(
      0.192450089729875,
      14,
    );
    expect(evaluateFormula("=BETA.DIST(0.4,2,3,TRUE)", workbook).value).toBeCloseTo(
      0.524800000000001,
      14,
    );
    expect(evaluateFormula("=F.DIST(1,2,3,TRUE)", workbook).value).toBeCloseTo(
      0.535241998455113,
      14,
    );
    expect(evaluateFormula("=T.DIST.RT(1.812468,10)", workbook).value).toBeCloseTo(0.05, 5);
    expect(evaluateFormula("=T.DIST.2T(1.812468,10)", workbook).value).toBeCloseTo(0.1, 4);
    expect(evaluateFormula("=T.INV.2T(0.1,10)", workbook).value).toBeCloseTo(1.812461, 5);
    expect(evaluateFormula("=T.INV(0.0000000001,2)", workbook).value).toBeCloseTo(-70710.678, 2);
    expect(evaluateFormula("=T.INV.2T(0.0000000001,2)", workbook).value).toBeCloseTo(100000, 2);
  });

  it("supports compatibility aliases and fixed-currency formatting", () => {
    expect(evaluateFormula("=LEGACY.CHIDIST(2,1)", workbook).value).toBe(
      evaluateFormula("=CHIDIST(2,1)", workbook).value,
    );
    expect(evaluateFormula("=LEGACY.TDIST(1,2,2)", workbook).value).toBeCloseTo(
      0.422649730810374,
      12,
    );
    expect(evaluateFormula("=USDOLLAR(1234.5,1)", workbook).value).toBe("$1,234.5");
  });

  it("uses host text encodings for byte-oriented functions", () => {
    const encodingWorkbook: FormulaWorkbookContext = {
      ...workbook,
      textEncoding: {
        codePage: 936,
        encode: (text) => {
          if (text === "中") return new Uint8Array([0xd6, 0xd0]);
          return new TextEncoder().encode(text);
        },
        decode: (bytes) => (bytes.length === 2 ? "中" : String.fromCharCode(bytes[0]!)),
      },
    };
    expect(evaluateFormula('=LENB("A中")', encodingWorkbook).value).toBe(3);
    expect(evaluateFormula('=LEFTB("A中",1)', encodingWorkbook).value).toBe("A");
    expect(evaluateFormula('=CODE("中")', encodingWorkbook).value).toBe(214);
    expect(evaluateFormula("=CHAR(65)", encodingWorkbook).value).toBe("A");
  });

  it("can return structured runtime errors", () => {
    const result = evaluateFormula("=1/0", workbook, { errors: "object" });
    expect(result.value).toBeInstanceOf(formulaPublic.FormulaError);
    expect(evaluateFormula("=1/0", workbook).value).toBe("#DIV/0!");
  });

  it("extracts regular expressions and reports invalid patterns", () => {
    expect(evaluateFormula('=REGEXEXTRACT("abc123","[0-9]+")', workbook).value).toBe("123");
    expect(evaluateFormula('=REGEXEXTRACT("abc","(")', workbook).value).toBe("#VALUE!");
  });

  it("parses Excel this-row and OpenFormula sheet-qualified selectors", () => {
    expect(parseFormula("=Sales[@Amount]")).toMatchObject({
      reference: {
        table: {
          item: "this-row",
          selectors: [{ kind: "column", column: "Amount" }],
        },
      },
    });
    expect(parseFormula("=Sales[[#This Row],[Amount]]")).toMatchObject({
      reference: {
        table: {
          item: "this-row",
          selectors: [{ kind: "column", column: "Amount" }],
        },
      },
    });
    expect(parseFormula("=[Sheet1.A1:Sheet2.B2]", { dialect: "openformula" })).toMatchObject({
      reference: {
        sheets: { names: ["Sheet1", "Sheet2"], range: true },
        areas: [
          {
            fromSheet: "Sheet1",
            toSheet: "Sheet2",
            start: { row: 1, column: 1 },
            end: { row: 2, column: 2 },
          },
        ],
      },
    });
  });

  it("expands table dependencies and local external workbook identities", () => {
    const tableGraph = buildDependencyGraph(
      [
        {
          id: "formula",
          ast: parseFormula("=Sales[Alpha]"),
          cell: { sheet: "Data", row: 1, column: 3 },
        },
        { id: "first", ast: parseFormula("=1"), cell: { sheet: "Data", row: 2, column: 1 } },
        { id: "second", ast: parseFormula("=2"), cell: { sheet: "Data", row: 3, column: 1 } },
      ],
      {
        getTableName: (name) =>
          name === "Sales"
            ? {
                sheet: "Data",
                range: { firstRow: 2, firstColumn: 1, lastRow: 3, lastColumn: 2 },
                columnNames: ["Alpha", "Beta"],
              }
            : undefined,
      },
    );
    expect([...tableGraph.edges.get("formula")!].sort()).toEqual(["first", "second"]);

    const bareTableGraph = buildDependencyGraph(
      [
        {
          id: "formula",
          ast: parseFormula("=Sales"),
          cell: { sheet: "Data", row: 1, column: 3 },
        },
        { id: "first", ast: parseFormula("=1"), cell: { sheet: "Data", row: 2, column: 1 } },
        { id: "second", ast: parseFormula("=2"), cell: { sheet: "Data", row: 3, column: 1 } },
      ],
      {
        getTableName: (name) =>
          name === "Sales"
            ? {
                sheet: "Data",
                range: { firstRow: 2, firstColumn: 1, lastRow: 3, lastColumn: 2 },
                columnNames: ["Alpha", "Beta"],
              }
            : undefined,
      },
    );
    expect([...bareTableGraph.edges.get("formula")!].sort()).toEqual(["first", "second"]);
    expect(collectDirtyFormulas(bareTableGraph, ["first"]).has("formula")).toBe(true);

    const localGraph = buildDependencyGraph([
      {
        id: "formula",
        ast: parseFormula("=A1"),
        cell: { workbook: "Book", sheet: "Sheet1", row: 1, column: 2 },
      },
      {
        id: "local",
        ast: parseFormula("=1"),
        cell: { workbook: "Book", sheet: "Sheet1", row: 1, column: 1 },
      },
      {
        id: "other",
        ast: parseFormula("=2"),
        cell: { sheet: "Sheet1", row: 1, column: 1 },
      },
    ]);
    expect([...localGraph.edges.get("formula")!]).toEqual(["local"]);
  });

  it("does not expand nonadjacent table columns into their gap", () => {
    const graph = buildDependencyGraph(
      [
        {
          id: "formula",
          ast: parseFormula("=SUM(Sales[[Alpha],[Gamma]])"),
          cell: { sheet: "Data", row: 1, column: 1 },
        },
        { id: "alpha", ast: parseFormula("=1"), cell: { sheet: "Data", row: 2, column: 5 } },
        { id: "beta", ast: parseFormula("=2"), cell: { sheet: "Data", row: 2, column: 2 } },
        { id: "gamma", ast: parseFormula("=3"), cell: { sheet: "Data", row: 2, column: 7 } },
      ],
      {
        getTableName: (name) =>
          name === "Sales"
            ? {
                sheet: "Data",
                range: { firstRow: 2, firstColumn: 5, lastRow: 3, lastColumn: 7 },
                columnNames: ["Alpha", "Beta", "Gamma"],
              }
            : undefined,
      },
    );
    expect([...graph.edges.get("formula")!].sort()).toEqual(["alpha", "gamma"]);
  });

  it("narrows table item dependencies to their represented rows", () => {
    const tableInfo = {
      sheet: "Data",
      range: { firstRow: 1, firstColumn: 5, lastRow: 4, lastColumn: 7 },
      columnNames: ["Alpha", "Beta", "Gamma"],
      hasHeaders: true,
      hasTotals: true,
    };
    const headerGraph = buildDependencyGraph(
      [
        {
          id: "formula",
          ast: parseFormula("=Sales[[#Headers],[Alpha]]"),
          cell: { sheet: "Data", row: 3, column: 5 },
        },
        { id: "header", ast: parseFormula("=1"), cell: { sheet: "Data", row: 1, column: 5 } },
        { id: "body", ast: parseFormula("=2"), cell: { sheet: "Data", row: 3, column: 5 } },
        { id: "total", ast: parseFormula("=3"), cell: { sheet: "Data", row: 4, column: 5 } },
      ],
      { getTableName: (name) => (name === "Sales" ? tableInfo : undefined) },
    );
    expect([...headerGraph.edges.get("formula")!]).toEqual(["header"]);
    expect(detectFormulaCycles(headerGraph)).toEqual([]);

    const thisRowGraph = buildDependencyGraph(
      [
        {
          id: "formula",
          ast: parseFormula("=Sales[[#This Row],[Alpha]]"),
          cell: { sheet: "Data", row: 3, column: 6 },
        },
        { id: "alpha", ast: parseFormula("=1"), cell: { sheet: "Data", row: 3, column: 5 } },
        { id: "otherRow", ast: parseFormula("=2"), cell: { sheet: "Data", row: 2, column: 5 } },
      ],
      { getTableName: (name) => (name === "Sales" ? tableInfo : undefined) },
    );
    expect([...thisRowGraph.edges.get("formula")!]).toEqual(["alpha"]);
    expect(detectFormulaCycles(thisRowGraph)).toEqual([]);

    const bareThisRowGraph = buildDependencyGraph(
      [
        {
          id: "formula",
          ast: parseFormula("=@Sales"),
          cell: { sheet: "Data", row: 3, column: 8 },
        },
        { id: "alpha", ast: parseFormula("=1"), cell: { sheet: "Data", row: 3, column: 5 } },
        { id: "otherRow", ast: parseFormula("=2"), cell: { sheet: "Data", row: 2, column: 5 } },
      ],
      { getTableName: (name) => (name === "Sales" ? tableInfo : undefined) },
    );
    expect([...bareThisRowGraph.edges.get("formula")!]).toEqual(["alpha"]);
    expect(detectFormulaCycles(bareThisRowGraph)).toEqual([]);

    for (const row of [tableInfo.range.firstRow, tableInfo.range.lastRow]) {
      const boundaryGraph = buildDependencyGraph(
        [
          {
            id: "formula",
            ast: parseFormula("=Sales[[#This Row],[Alpha]]"),
            cell: { sheet: "Data", row, column: 6 },
          },
        ],
        { getTableName: (name) => (name === "Sales" ? tableInfo : undefined) },
      );
      expect(boundaryGraph.edges.get("formula")?.size).toBe(0);
    }

    const emptyBodyGraph = buildDependencyGraph(
      [
        {
          id: "formula",
          ast: parseFormula("=Sales[#Data]"),
          cell: { sheet: "Data", row: 1, column: 1 },
        },
      ],
      {
        getTableName: (name) =>
          name === "Sales"
            ? {
                sheet: "Data",
                range: { firstRow: 1, firstColumn: 1, lastRow: 2, lastColumn: 2 },
                columnNames: ["Alpha", "Beta"],
                hasHeaders: true,
                hasTotals: true,
              }
            : undefined,
      },
    );
    expect(emptyBodyGraph.edges.get("formula")?.size).toBe(0);
  });

  it("keeps missing OFFSET arguments positional and records dynamic dependencies", () => {
    expect(evaluateFormula("=OFFSET(A1,1,1,,2)", workbook).value).toEqual([[22, 23]]);
    const offsetDependencies = collectDependencies(parseFormula("=OFFSET(A1,1,1,,2)"));
    expect(offsetDependencies.at(-1)?.target).toMatchObject({
      from: { row: 2, column: 2 },
      to: { row: 2, column: 3 },
    });
    const indirect = collectDependencies(parseFormula('=INDIRECT("R1C1",FALSE)'));
    expect(indirect.some(({ target }) => target.kind === "unknown")).toBe(true);
    expect(
      collectDependencies(parseFormula("=INDIRECT(A1)")).some(
        ({ target }) => target.kind === "unknown",
      ),
    ).toBe(true);
  });

  it("applies AGGREGATE nested aggregation only for options 0 through 3", () => {
    const nestedWorkbook: FormulaWorkbookContext = {
      ...workbook,
      getCell: (_sheet, row, column) =>
        row === 1 && column === 1
          ? { value: 11, formula: "=SUBTOTAL(9,B1)" }
          : { value: row * 10 + column },
    };
    expect(evaluateFormula("=AGGREGATE(9,0,A1:A2)", nestedWorkbook).value).toBe(21);
    expect(evaluateFormula("=AGGREGATE(9,4,A1:A2)", nestedWorkbook).value).toBe(32);
  });

  it("normalizes non-finite matrix results", () => {
    expect(evaluateFormula("={1000,1000}^{1000,1000}", workbook).value).toEqual([
      ["#NUM!", "#NUM!"],
    ]);
  });

  it("uses table sheet identity separately from table id", () => {
    const tableWorkbook: FormulaWorkbookContext = {
      ...workbook,
      getTableName: (name) =>
        name === "Sales"
          ? {
              id: 17,
              sheetId: 1,
              range: { firstRow: 1, firstColumn: 1, lastRow: 1, lastColumn: 1 },
            }
          : undefined,
    };
    expect(
      evaluateFormula("=Sales[#This Row]", tableWorkbook, { sheet: 1, row: 1, column: 1 }).value,
    ).toBe(11);
    expect(
      evaluateFormula("=Sales[#This Row]", tableWorkbook, { sheet: 0, row: 1, column: 1 }).value,
    ).toBe("#VALUE!");
  });
});

describe("formula semantic regressions", () => {
  const workbook: FormulaWorkbookContext = {
    getSheetId: (name) => (name === "Sheet1" ? 0 : undefined),
    getSheetName: (id) => (id === 0 ? "Sheet1" : undefined),
    getCell: (_sheet, row, column) => ({ value: row * 10 + column }),
    getDefinedFormula: () => undefined,
    getTableName: () => undefined,
    getRowCount: () => 100,
    getColumnCount: () => 100,
  };

  it("keeps matrix inversion and determinant scalar-valued", () => {
    expect(evaluateFormula("=MDETERM({1,2;3,4})", workbook).value).toBe(-2);
    expect(evaluateFormula("=MINVERSE({4,7;2,6})", workbook).value).toEqual([
      [expect.closeTo(0.6, 15), expect.closeTo(-0.7, 15)],
      [expect.closeTo(-0.2, 15), expect.closeTo(0.4, 15)],
    ]);
  });

  it("returns wildcard and byte-search starts in the caller's unit", () => {
    expect(evaluateFormula('=SEARCH("bc","abc")', workbook).value).toBe(2);
    expect(evaluateFormula('=SEARCH("b","abcb",2)', workbook).value).toBe(2);
    expect(evaluateFormula('=SEARCHB("b","a𐐷𐐷b")', workbook).value).toBe(6);
    expect(evaluateFormula('=FINDB("b","a𐐷𐐷b")', workbook).value).toBe(6);
  });

  it("passes AST references to SQL hosts without evaluating them", () => {
    const requests: unknown[] = [];
    const sqlWorkbook: FormulaWorkbookContext = {
      ...workbook,
      executeSql: (request) => {
        requests.push(request);
        return "done";
      },
    };
    expect(
      evaluateFormula(
        '=SQL.REQUEST("connection",Sheet1!A1,2,"select",TRUE,FALSE,10,20,30)',
        sqlWorkbook,
      ).value,
    ).toBe("done");
    expect(requests).toEqual([
      {
        connectionString: "connection",
        driverPrompt: 2,
        query: "select",
        outputReference: "Sheet1!A1",
        columnNames: true,
        rowNames: false,
        maxRows: 10,
        maxChanges: 20,
        timeout: 30,
      },
    ]);
    requests.length = 0;
    expect(evaluateFormula('=SQL.REQUEST("c",,2,"q")', sqlWorkbook).value).toBe("done");
    expect(requests).toEqual([{ connectionString: "c", driverPrompt: 2, query: "q" }]);
  });

  it("validates dynamic-array dimensions and date iteration bounds", () => {
    expect(evaluateFormula('=EXPAND({1,2},"x")', workbook).value).toBe("#VALUE!");
    expect(evaluateFormula("=RANDARRAY(0)", workbook).value).toBe("#VALUE!");
    expect(evaluateFormula("=SEQUENCE(1.5)", workbook).value).toBe("#VALUE!");
    expect(evaluateFormula("=MAKEARRAY(1.5,1,LAMBDA(row,column,1))", workbook).value).toBe(
      "#VALUE!",
    );
    expect(evaluateFormula("=MAKEARRAY(9E+15,2,LAMBDA(row,column,1))", workbook).value).toBe(
      "#VALUE!",
    );
    expect(evaluateFormula("=SEQUENCE(1415,1415)", workbook).value).toBe("#VALUE!");
    expect(evaluateFormula("=MUNIT(1415)", workbook).value).toBe("#VALUE!");
    expect(evaluateFormula("=MAKEARRAY(1415,1415,LAMBDA(row,column,1))", workbook).value).toBe(
      "#VALUE!",
    );
    expect(evaluateFormula("=MMULT(SEQUENCE(2000,1),SEQUENCE(1,2000))", workbook).value).toBe(
      "#VALUE!",
    );
    expect(evaluateFormula("=SEQUENCE(1000001):SEQUENCE(1000001)", workbook).value).toBe("#VALUE!");
    expect(evaluateFormula("=WRAPROWS(1,2000001)", workbook).value).toBe("#VALUE!");
    expect(evaluateFormula("=WRAPCOLS(1,2000001)", workbook).value).toBe("#VALUE!");
    expect(evaluateFormula("=WRAPROWS(1,9007199254740991)", workbook).value).toBe("#VALUE!");
    expect(evaluateFormula('=TEXTSPLIT(REPT(",",2000001),",")', workbook).value).toBe("#VALUE!");
    expect(evaluateFormula("=SEQUENCE(1500)*SEQUENCE(1,1501)", workbook).value).toBe("#VALUE!");
    expect(
      evaluateFormula("=GROUPBY(SEQUENCE(1000001),SEQUENCE(1000001),SUM)", workbook).value,
    ).toBe("#VALUE!");
    expect(
      evaluateFormula(
        "=PIVOTBY(SEQUENCE(1000,1001),SEQUENCE(1000,1001),SEQUENCE(1000),SUM)",
        workbook,
      ).value,
    ).toBe("#VALUE!");
    expect(evaluateFormula("=ROUND(SEQUENCE(1500),SEQUENCE(1,1501))", workbook).value).toBe(
      "#VALUE!",
    );
    expect(evaluateFormula("=IF(SEQUENCE(1500),SEQUENCE(1,1501),0)", workbook).value).toBe(
      "#VALUE!",
    );
    expect(evaluateFormula("=IFERROR(SEQUENCE(1500)/0,SEQUENCE(1,1501))", workbook).value).toBe(
      "#VALUE!",
    );
    expect(evaluateFormula("=EXPAND({1,2},2)", workbook).value).toEqual([
      [1, 2],
      ["#N/A", "#N/A"],
    ]);
    expect(evaluateFormula("=WRAPROWS({1;2;3},0)", workbook).value).toBe("#VALUE!");
    expect(evaluateFormula("=WRAPROWS({1;2;3},2)", workbook).value).toEqual([
      [1, 2],
      [3, "#N/A"],
    ]);
    expect(evaluateFormula("=VSTACK({1,2},{3})", workbook).value).toEqual([
      [1, 2],
      [3, "#N/A"],
    ]);
    expect(evaluateFormula("=HSTACK({1;2},{3})", workbook).value).toEqual([
      [1, 3],
      [2, "#N/A"],
    ]);
    expect(evaluateFormula("=NETWORKDAYS(-Infinity,1)", workbook).value).toBe("#VALUE!");
    expect(evaluateFormula("=NETWORKDAYS(9E+99,1)", workbook).value).toBe("#NUM!");
    expect(evaluateFormula("=NETWORKDAYS.INTL(1,Infinity)", workbook).value).toBe("#VALUE!");
    expect(evaluateFormula("=NETWORKDAYS.INTL(1,9E+99)", workbook).value).toBe("#NUM!");
    expect(evaluateFormula("=NETWORKDAYS(3,1)", workbook).value).toBe(-2);
    expect(evaluateFormula("=DAYS360(DATE(2026,1,31),DATE(2026,3,1),1)", workbook).value).toBe(31);
    expect(evaluateFormula("=YEAR(-1)", workbook).value).toBe("#NUM!");
    expect(evaluateFormula("=DAY(2958466)", workbook).value).toBe("#NUM!");
    expect(evaluateFormula("=DATE(10000,1,1)", workbook).value).toBe("#NUM!");
    expect(evaluateFormula("=EDATE(DATE(9999,12,31),1)", workbook).value).toBe("#NUM!");
    expect(
      evaluateFormula("=EDATE(DATE(9999,12,31),1)", workbook, { dateSystem: "1904" }).value,
    ).toBe("#NUM!");
    expect(evaluateFormula("=DAYS(9E+99,0)", workbook).value).toBe("#NUM!");
    const infiniteCell: FormulaWorkbookContext = {
      ...workbook,
      getCell: () => ({ value: Number.POSITIVE_INFINITY }),
    };
    expect(evaluateFormula("=NETWORKDAYS(A1,B1)", infiniteCell).value).toBe("#VALUE!");
    expect(evaluateFormula('=SEARCHB("x","abc")', workbook).value).toBe("#VALUE!");
  });

  it("normalizes regex and character-code boundary errors", () => {
    expect(evaluateFormula('=REGEXREPLACE("a","(","x")', workbook).value).toBe("#VALUE!");
    expect(evaluateFormula('=REGEXTEST("a","(")', workbook).value).toBe("#VALUE!");
    expect(evaluateFormula("=CHAR(0)", workbook).value).toBe("#VALUE!");
    expect(evaluateFormula("=CHAR(1.5)", workbook).value).toBe("#VALUE!");
    expect(evaluateFormula("=UNICHAR(0)", workbook).value).toBe("#VALUE!");
    expect(evaluateFormula("=UNICHAR(1114112)", workbook).value).toBe("#VALUE!");
  });

  it("requires repetition for modes and uses exact grouping output sizes", () => {
    expect(evaluateFormula("=MODE({1,2})", workbook).value).toBe("#N/A");
    expect(evaluateFormula("=MODE.MULT({1,2})", workbook).value).toBe("#N/A");
    expect(
      evaluateFormula("=GROUPBY(SEQUENCE(1000001,1,1,0),SEQUENCE(1000001,1,1,0),SUM)", workbook)
        .value,
    ).toEqual([[1, 1000001]]);
    expect(
      evaluateFormula(
        "=PIVOTBY(SEQUENCE(1000001,1,1,0),SEQUENCE(1000001,1,1,0),SEQUENCE(1000001,1,1,0),SUM)",
        workbook,
      ).value,
    ).toEqual([[1, 1000001]]);
  });

  it("preserves grouping filters and per-sheet whole-reference dimensions", () => {
    expect(
      evaluateFormula('=GROUPBY({"A";"B";"A"},{1;2;3},SUM,0,0,0,{TRUE;FALSE;TRUE})', workbook)
        .value,
    ).toEqual([["A", 4]]);
    expect(
      evaluateFormula(
        '=PIVOTBY({"A";"A";"A"},{"X";"X";"X"},{1;2;3},SUM,0,0,0,0,0,{FALSE;TRUE;TRUE})',
        workbook,
      ).value,
    ).toEqual([["A", 5]]);
    expect(evaluateFormula("=VALUES(SEQUENCE(1000001),SEQUENCE(1000001))", workbook).value).toBe(
      "#VALUE!",
    );
    expect(() => parseFormula("=A1048577:B1")).toThrow();
    expect(() => parseFormula("=A1048577#")).toThrow();

    const dimensionalWorkbook: FormulaWorkbookContext = {
      ...workbook,
      getSheetId: (name) => (name === "Sheet1" ? 0 : name === "Sheet2" ? 1 : undefined),
      getSheetName: (id) => (id === 0 ? "Sheet1" : id === 1 ? "Sheet2" : undefined),
      getCell: (sheet, row) => (row <= sheet + 2 ? { value: 1 } : undefined),
      getRowCount: (sheet) => (sheet === 0 ? 2 : 3),
      getColumnCount: () => 1,
    };
    expect(evaluateFormula("=COUNT(Sheet1:Sheet2!A:A)", dimensionalWorkbook).value).toBe(5);
    expect(evaluateFormula("=@Missing!A1", dimensionalWorkbook).value).toBe("#REF!");
  });

  it("freezes the catalog and exposes controlled AST extension", () => {
    expect(Object.isFrozen(formulaFunctions)).toBe(true);
    expect(Object.isFrozen(formulaFunctions[0])).toBe(true);
    expect(formulaPublic.setFunctionAstEvaluator).toBeTypeOf("function");
    expect(formulaPublic.clearFunctionAstEvaluator).toBeTypeOf("function");
    const trueFunction = formulaPublic.getFunction("TRUE")!;
    formulaPublic.setFunctionAstEvaluator?.("TRUE", () => 42);
    expect(evaluateFormula("=TRUE()", workbook).value).toBe(42);
    formulaPublic.clearFunctionAstEvaluator?.("TRUE");
    expect(evaluateFormula("=TRUE()", workbook).value).toBe(true);
    expect(trueFunction).toBeDefined();
  });

  it("reports excessive parser nesting structurally", () => {
    const source = `${"(".repeat(10000)}1${")".repeat(10000)}`;
    expect(() => parseFormula(source)).toThrow(formulaPublic.FormulaParseError);
  });
});
