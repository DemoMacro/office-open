import { type FormulaAst, traverseFormula } from "./ast";
import type { FormulaDialect } from "./ast";
import { parseFormula } from "./parser";
import { formatCell, formatReferenceArea, shiftCell, shiftRangeEdge } from "./references";
import { stringifyFormula } from "./stringify";

export interface FormulaShiftOptions {
  rows?: number;
  columns?: number;
  dialect?: FormulaDialect;
}

export function shiftFormula(formula: string, options: FormulaShiftOptions = {}): string {
  const ast = parseFormula(
    formula,
    options.dialect === undefined ? {} : { dialect: options.dialect },
  );
  traverseFormula(ast, (node) => {
    if (node.kind !== "reference" || !node.reference) return;
    const reference = { ...node.reference };
    if (reference.areas) {
      reference.areas = reference.areas.map((area) => ({
        ...area,
        start: shiftRangeEdge(area.start, options.rows ?? 0, options.columns ?? 0),
        ...(area.end
          ? { end: shiftRangeEdge(area.end, options.rows ?? 0, options.columns ?? 0) }
          : {}),
      }));
      node.reference = reference;
      node.raw = `${reference.areas
        .map((area) => formatReferenceArea(area, options.dialect ?? ast.dialect ?? "excel"))
        .join(",")}${reference.spill ? "#" : ""}`;
      return;
    }
    if (!reference.start) return;
    const start = shiftCell(reference.start, options.rows ?? 0, options.columns ?? 0);
    const end = reference.end
      ? shiftCell(reference.end, options.rows ?? 0, options.columns ?? 0)
      : undefined;
    node.reference = { ...reference, start, end };
    node.raw = `${end ? `${formatCell(start)}:${formatCell(end)}` : formatCell(start)}${
      reference.spill ? "#" : ""
    }`;
  });
  return stringifyFormula(ast);
}

export function shiftFormulaAst(ast: FormulaAst, options: FormulaShiftOptions = {}): FormulaAst {
  return parseFormula(shiftFormula(ast.raw, { ...options, dialect: ast.dialect ?? "excel" }), {
    dialect: ast.dialect ?? "excel",
  });
}
