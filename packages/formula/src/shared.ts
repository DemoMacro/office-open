import type { FormulaSharedGroup } from "./context";
import { parseFormula } from "./parser";
import { shiftFormulaAst } from "./rewrite";

export interface FormulaPosition {
  readonly row: number;
  readonly column: number;
}

export function resolveSharedFormula(
  group: FormulaSharedGroup,
  slave: FormulaPosition,
  dialect: "excel" | "openformula" = "excel",
): ReturnType<typeof parseFormula> {
  const master = parseFormula("=" + group.formula, { dialect });
  return shiftFormulaAst(master, {
    rows: slave.row - group.masterRow,
    columns: slave.column - group.masterColumn,
  });
}
