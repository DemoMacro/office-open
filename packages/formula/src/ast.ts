import type { FormulaErrorValue } from "./value";

export type FormulaDialect = "excel" | "openformula";

export interface FormulaSpan {
  start: number;
  end: number;
}

export type FormulaNodeKind =
  | "array"
  | "binary"
  | "boolean"
  | "error"
  | "group"
  | "missing"
  | "name"
  | "number"
  | "reference"
  | "string"
  | "unary"
  | "postfix"
  | "call"
  | "application";

export interface FormulaCellCoordinate {
  row: number;
  column: number;
  absoluteRow: boolean;
  absoluteColumn: boolean;
}

export type FormulaRangeEdge =
  | FormulaCellCoordinate
  | {
      kind: "whole-column";
      column: number;
      absolute: boolean;
    }
  | {
      kind: "whole-row";
      row: number;
      absolute: boolean;
    };

export interface FormulaWorkbookSource {
  index?: number;
  path?: string;
  name?: string;
}

export type FormulaWorkbookIdentity = string | FormulaWorkbookSource;

export interface FormulaSheetSelector {
  names: string[];
  range: boolean;
}

export type FormulaTableColumnSelector =
  | { kind: "column"; column: string }
  | { kind: "range"; start: string; end: string };

export interface FormulaTableSelector {
  table: string;
  item?: "all" | "data" | "headers" | "totals" | "this-row";
  selectors: FormulaTableColumnSelector[];
}

export interface FormulaReferenceInfo {
  workbook?: FormulaWorkbookSource;
  sheets?: FormulaSheetSelector;
  areas?: FormulaReferenceArea[];
  start?: FormulaCellCoordinate;
  end?: FormulaCellCoordinate;
  spill?: boolean;
  implicitIntersection?: boolean;
  table?: FormulaTableSelector;
}

export interface FormulaReferenceArea {
  workbook?: FormulaWorkbookSource;
  fromSheet?: string;
  toSheet?: string;
  start: FormulaRangeEdge;
  end?: FormulaRangeEdge;
}

interface FormulaAstBase extends FormulaSpan {
  raw: string;
  prefix?: "=";
  dialect?: FormulaDialect;
}

export interface FormulaNumberAst extends FormulaAstBase {
  kind: "number";
  value: number;
}

export interface FormulaStringAst extends FormulaAstBase {
  kind: "string";
  value: string;
}

export interface FormulaBooleanAst extends FormulaAstBase {
  kind: "boolean";
  value: boolean;
}

export interface FormulaErrorAst extends FormulaAstBase {
  kind: "error";
  value: FormulaErrorValue;
}

export interface FormulaGroupAst extends FormulaAstBase {
  kind: "group";
  operand: FormulaAst;
}

export interface FormulaReferenceAst extends FormulaAstBase {
  kind: "reference";
  reference?: FormulaReferenceInfo;
}

export interface FormulaNameAst extends FormulaAstBase {
  kind: "name";
  name: string;
  implicitIntersection?: boolean;
}

export interface FormulaMissingAst extends FormulaAstBase {
  kind: "missing";
}

export interface FormulaArrayAst extends FormulaAstBase {
  kind: "array";
  rows: FormulaAst[][];
}

export interface FormulaUnaryAst extends FormulaAstBase {
  kind: "unary";
  operator: "+" | "-";
  operand: FormulaAst;
}

export interface FormulaPostfixAst extends FormulaAstBase {
  kind: "postfix";
  operator: "%";
  operand: FormulaAst;
}

export type FormulaBinaryOperator =
  | "+"
  | "-"
  | "*"
  | "/"
  | "^"
  | "&"
  | "="
  | "<>"
  | "<"
  | "<="
  | ">"
  | ">="
  | ":"
  | "!"
  | " "
  | "~"
  | ",";

export interface FormulaBinaryAst extends FormulaAstBase {
  kind: "binary";
  operator: FormulaBinaryOperator;
  left: FormulaAst;
  right: FormulaAst;
}

export interface FormulaCallAst extends FormulaAstBase {
  kind: "call";
  name: string;
  args: FormulaAst[];
}

export interface FormulaApplicationAst extends FormulaAstBase {
  kind: "application";
  callee: FormulaAst;
  args: FormulaAst[];
}

export type FormulaAst =
  | FormulaArrayAst
  | FormulaApplicationAst
  | FormulaBinaryAst
  | FormulaBooleanAst
  | FormulaCallAst
  | FormulaErrorAst
  | FormulaGroupAst
  | FormulaMissingAst
  | FormulaNameAst
  | FormulaNumberAst
  | FormulaPostfixAst
  | FormulaReferenceAst
  | FormulaStringAst
  | FormulaUnaryAst;

export interface FormulaParseOptions {
  dialect?: FormulaDialect;
}

export function traverseFormula(ast: FormulaAst, visit: (node: FormulaAst) => void): void {
  visit(ast);
  if (ast.kind === "array") {
    for (const row of ast.rows) for (const child of row) traverseFormula(child, visit);
    return;
  }
  if (ast.kind === "group" || ast.kind === "unary" || ast.kind === "postfix") {
    traverseFormula(ast.operand, visit);
    return;
  }
  if (ast.kind === "binary") {
    traverseFormula(ast.left, visit);
    traverseFormula(ast.right, visit);
    return;
  }
  if (ast.kind === "call" || ast.kind === "application") {
    if (ast.kind === "application") traverseFormula(ast.callee, visit);
    for (const child of ast.args) traverseFormula(child, visit);
  }
}
