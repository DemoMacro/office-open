import {
  buildDependencyGraph,
  collectDependencies,
  evaluateFormulaAst,
  FormulaError,
  isFormulaErrorToken,
  parseFormula,
  resolveSharedFormula,
  topologicalFormulaOrder,
  type FormulaCellSnapshot,
  type FormulaDialect,
  type FormulaEvaluationOptions,
  type FormulaMatrix,
  type FormulaScalar,
  type FormulaValue,
  type FormulaWorkbookContext,
} from "@office-open/formula";
import type {
  CellOptions,
  FormulaOptions,
  RowOptions,
  TableOptions,
  WorksheetOptions,
} from "@office-open/xlsx";
import type { WorkbookOptions } from "@office-open/xlsx";

type CanonicalWorkbookOptions = WorkbookOptions;

export type WorkbookFormulaDialect = FormulaDialect;

export interface WorkbookFormulaPosition {
  sheet: number;
  row: number;
  column: number;
}

export interface WorkbookFormulaCalculation {
  readonly sheet: string;
  readonly row: number;
  readonly column: number;
  readonly reference: string;
  readonly formula: string;
  readonly type: Required<FormulaOptions>["type"];
  readonly value: FormulaScalar;
  readonly previousValue: FormulaScalar;
}

export interface WorkbookFormulaRecalculation {
  readonly results: readonly WorkbookFormulaCalculation[];
  readonly failures: readonly WorkbookFormulaFailure[];
}

export interface WorkbookFormulaFailure {
  readonly sheet: string;
  readonly reference: string;
  readonly formula: string;
  readonly reason: string;
}

interface WorkbookCell {
  readonly sheetIndex: number;
  readonly row: number;
  readonly column: number;
  readonly reference: string;
  readonly cell: CellOptions;
}

interface ResolvedWorkbookFormula {
  readonly id: string;
  readonly source: WorkbookCell;
  readonly formula: string;
  readonly type: Required<FormulaOptions>["type"];
  readonly ast: ReturnType<typeof parseFormula>;
  readonly followers?: readonly WorkbookCell[];
}

export interface WorkbookFormulaIssue {
  readonly location: string;
  readonly formula: string;
  readonly message: string;
}

const COLUMN_PATTERN = /^(\$?)([A-Z]{1,3})(\$?)(\d{1,7})$/u;
const ERROR_PREFIX = "#";

function normalizeOdfNamedRange(value: string): string {
  const edges = value.split(":");
  if (
    edges.length > 2 ||
    !edges.every((edge, index) =>
      index === 0 ? /^[^!]+\.[$A-Za-z0-9]+$/u.test(edge) : /^[$A-Za-z0-9]+$/u.test(edge),
    )
  )
    return value;
  const sheet = edges[0]!.slice(0, edges[0]!.indexOf(".")).replace(/^\$/u, "");
  const range = edges
    .map((edge, index) => `${index === 0 ? sheet : ""}.${edge.slice(edge.indexOf(".") + 1)}`)
    .join(":");
  return `[${range}]`;
}

export function createWorkbookFormulaContext(
  options: CanonicalWorkbookOptions,
  dialect: WorkbookFormulaDialect = "excel",
): FormulaWorkbookContext {
  const sheets = options.worksheets ?? [];
  const names = sheets.map((sheet, index) => sheet.name ?? `Sheet${index + 1}`);
  const cells = collectWorkbookCells(sheets);
  const cellIndex = new Map(
    cells.map((entry) => [cellKey(entry.sheetIndex, entry.row, entry.column), entry]),
  );
  const formulaByCell = new Map<string, string>();
  for (const formula of collectResolvedFormulas(options, dialect)) {
    const cells =
      formula.type === "array" ? [formula.source] : [formula.source, ...(formula.followers ?? [])];
    for (const cell of cells)
      formulaByCell.set(cellKey(cell.sheetIndex, cell.row, cell.column), formula.formula);
  }
  const tables = new Map(
    sheets.flatMap((sheet, sheetIndex) =>
      (sheet.tables ?? []).map((table) => [tableName(table), { sheetIndex, table }]),
    ),
  );

  return {
    getSheetId: (name) => {
      const index = names.findIndex((candidate) => candidate.toLowerCase() === name.toLowerCase());
      return index === -1 ? undefined : index;
    },
    getSheetName: (id) => names[id],
    getCell: (sheet, row, column) => {
      const entry = cellIndex.get(cellKey(sheet, row, column));
      if (!entry) return undefined;
      const snapshot: FormulaCellSnapshot = {
        formula: formulaByCell.get(cellKey(sheet, row, column)),
      };
      if (entry.cell.error !== undefined) {
        snapshot.value = normalizeErrorToken(entry.cell.error);
        snapshot.error = true;
      } else snapshot.value = toFormulaScalar(entry.cell.value);
      return snapshot;
    },
    getDefinedFormula: (name, sheetId) => {
      const key = name.toLowerCase();
      const defined = options.definedNames?.filter(
        (entry) =>
          entry.name.toLowerCase() === key &&
          (entry.localSheetId === undefined || entry.localSheetId === sheetId),
      );
      const value =
        defined?.find((entry) => entry.localSheetId === sheetId)?.value ?? defined?.[0]?.value;
      return value === undefined ? undefined : normalizeOdfNamedRange(value);
    },
    getTableName: (name) => {
      const found = tables.get(name.toLowerCase());
      if (!found) return undefined;
      const range = parseRange(found.table.ref);
      if (!range) return undefined;
      const hasHeaders = (found.table.headerRowCount ?? 1) > 0;
      const hasTotals = (found.table.totalsRowCount ?? 0) > 0;
      return {
        id: found.table.id ?? found.sheetIndex + 1,
        sheetId: found.sheetIndex,
        range,
        columnNames: found.table.columns.map((column) => column.name ?? column.uniqueName),
        hasHeaders,
        hasTotals,
      };
    },
    getRowCount: (sheet) => {
      const maximum = cells.reduce(
        (value, entry) => (entry.sheetIndex === sheet ? Math.max(value, entry.row) : value),
        0,
      );
      return maximum || 1_048_576;
    },
    getColumnCount: (sheet) => {
      const maximum = cells.reduce(
        (value, entry) => (entry.sheetIndex === sheet ? Math.max(value, entry.column) : value),
        0,
      );
      return maximum || 16_384;
    },
    isRowHidden: (sheet, row) =>
      sheets[sheet]?.rows?.find((entry) => actualRowNumber(entry) === row)?.hidden === true,
  };
}

export function evaluateWorkbookFormula(
  options: CanonicalWorkbookOptions,
  formula: string,
  position?: Partial<WorkbookFormulaPosition>,
  evaluationOptions: Omit<FormulaEvaluationOptions, "sheet" | "row" | "column"> = {},
): FormulaValue {
  return evaluateFormulaAst(
    parseWorkbookFormula(formula, evaluationOptions.dialect),
    createWorkbookFormulaContext(options, evaluationOptions.dialect),
    {
      ...evaluationOptions,
      sheet: position?.sheet ?? 0,
      row: position?.row ?? 1,
      column: position?.column ?? 1,
    },
  );
}

export function lintWorkbookFormulas(
  options: CanonicalWorkbookOptions,
  dialect: WorkbookFormulaDialect = "excel",
): readonly WorkbookFormulaIssue[] {
  const context = createWorkbookFormulaContext(options, dialect);
  const issues: WorkbookFormulaIssue[] = [];
  const check = (formula: string, location: string): void => {
    let ast;
    try {
      ast = parseWorkbookFormula(formula, dialect);
    } catch (error) {
      issues.push({
        location,
        formula,
        message: error instanceof Error ? error.message : String(error),
      });
      return;
    }
    for (const dependency of collectDependencies(ast)) {
      const message = dependencyIssue(dependency.target, context, options, dialect);
      if (message) issues.push({ location, formula, message });
    }
  };
  (options.worksheets ?? []).forEach((worksheet, sheetIndex) => {
    const name = worksheetName(options, sheetIndex);
    for (const row of worksheet.rows ?? [])
      for (const cell of row.cells ?? []) {
        const formula = formulaOptions(cell)?.formula;
        if (formula)
          check(
            formula,
            `${name}!${cell.reference ?? formatReference(actualRowNumber(row) ?? 1, 1)}`,
          );
      }
    for (const format of worksheet.conditionalFormats ?? [])
      for (const rule of format.rules ?? [])
        for (const formula of rule.formulas ?? [])
          check(formula, `${name} conditional format ${format.sqref}`);
    for (const validation of worksheet.dataValidations ?? [])
      for (const formula of [validation.formula1, validation.formula2])
        if (formula) check(formula, `${name} data validation ${validation.sqref}`);
  });
  return issues;
}

function dependencyIssue(
  target: ReturnType<typeof collectDependencies>[number]["target"],
  context: FormulaWorkbookContext,
  options: CanonicalWorkbookOptions,
  dialect: WorkbookFormulaDialect,
): string | undefined {
  if (target.workbook !== undefined) return undefined;
  if (target.kind === "name") {
    return context.getDefinedFormula(target.name ?? "") === undefined
      ? `references undefined name "${target.name ?? ""}"`
      : undefined;
  }
  if (target.kind === "table")
    return context.getTableName(target.name ?? "") === undefined
      ? `references missing table "${target.name ?? ""}"`
      : undefined;
  if (target.sheet === undefined) return undefined;
  const names = (options.worksheets ?? []).map(
    (worksheet, index) => worksheet.name ?? `Sheet${index + 1}`,
  );
  const from = context.getSheetId(target.sheet);
  const to = target.toSheet === undefined ? from : context.getSheetId(target.toSheet);
  if (from === undefined)
    return `references missing sheet "${target.sheet}" (sheets: ${names.join(", ")})`;
  if (to === undefined)
    return `references missing sheet "${target.toSheet ?? ""}" (sheets: ${names.join(", ")})`;
  if (target.toSheet !== undefined) {
    const first = Math.min(from, to);
    const last = Math.max(from, to);
    const missing = names
      .slice(first, last + 1)
      .find((name) => context.getSheetId(name) === undefined);
    if (missing) return `references missing sheet "${missing}" (sheets: ${names.join(", ")})`;
  }
  if (dialect === "openformula" && target.sheet.startsWith("[") && target.sheet.endsWith("]"))
    return undefined;
  return context.getSheetId(target.sheet) === undefined
    ? `references missing sheet "${target.sheet}"`
    : undefined;
}

export function calculateWorkbookFormulas(
  options: CanonicalWorkbookOptions,
  evaluationOptions: Pick<FormulaEvaluationOptions, "dateSystem" | "dialect" | "errors"> = {},
): WorkbookFormulaRecalculation {
  const formulas = collectResolvedFormulas(options, evaluationOptions.dialect);
  const context = createWorkbookFormulaContext(options, evaluationOptions.dialect);
  const graph = buildDependencyGraph(
    [
      ...formulas.map((formula) => ({
        id: formula.id,
        ast: formula.ast,
        cell: {
          sheet: worksheetName(options, formula.source.sheetIndex),
          row: formula.source.row,
          column: formula.source.column,
        },
      })),
      ...(options.definedNames ?? []).map((definedName, index) => ({
        id: `#defined-name:${index}`,
        name: definedName.name,
        ...(definedName.localSheetId === undefined
          ? {}
          : { scopeSheet: worksheetName(options, definedName.localSheetId) }),
        ast: parseWorkbookFormula(
          normalizeOdfNamedRange(definedName.value),
          evaluationOptions.dialect,
        ),
      })),
    ],
    {
      sheetOrder: (options.worksheets ?? []).map(
        (sheet, index) => sheet.name ?? `Sheet${index + 1}`,
      ),
      getTableName: (name) => context.getTableName(name),
    },
  );
  const formulaById = new Map(formulas.map((formula) => [formula.id, formula]));
  const results: WorkbookFormulaCalculation[] = [];
  const failures: WorkbookFormulaFailure[] = [];

  for (const id of topologicalFormulaOrder(graph)) {
    const formula = formulaById.get(id);
    if (!formula) continue;
    const previous = toFormulaScalar(formula.source.cell.value);
    try {
      const value = evaluateFormulaAst(formula.ast, context, {
        ...evaluationOptions,
        sheet: formula.source.sheetIndex,
        row: formula.source.row,
        column: formula.source.column,
      });
      const targets = formula.followers?.length ? formula.followers : [formula.source];
      distributeFormulaValue(targets, value, formula.type);
      const scalar = firstScalar(value);
      results.push({
        sheet: worksheetName(options, formula.source.sheetIndex),
        row: formula.source.row,
        column: formula.source.column,
        reference: formula.source.reference,
        formula: formula.formula,
        type: formula.type,
        value: scalar,
        previousValue: previous,
      });
    } catch (error) {
      failures.push({
        sheet: worksheetName(options, formula.source.sheetIndex),
        reference: formula.source.reference,
        formula: formula.formula,
        reason: error instanceof Error ? error.message : String(error),
      });
    }
  }
  return { results, failures };
}

function collectWorkbookCells(sheets: readonly WorksheetOptions[]): WorkbookCell[] {
  return sheets.flatMap((sheet, sheetIndex) =>
    (sheet.rows ?? []).flatMap((row) =>
      (row.cells ?? []).flatMap((cell) => {
        const position = parseReference(cell.reference ?? "");
        if (!position) return [];
        return [
          {
            sheetIndex,
            row: position.row,
            column: position.column,
            reference: cell.reference ?? formatReference(position.row, position.column),
            cell,
          },
        ];
      }),
    ),
  );
}

function collectResolvedFormulas(
  options: CanonicalWorkbookOptions,
  dialect: WorkbookFormulaDialect = "excel",
): ResolvedWorkbookFormula[] {
  const sheets = options.worksheets ?? [];
  const cells = collectWorkbookCells(sheets);
  const shared = new Map<string, WorkbookCell>();
  for (const entry of cells) {
    const formula = formulaOptions(entry.cell);
    if (formula?.type !== "shared" || !formula.formula || formula.sharedIndex === undefined)
      continue;
    shared.set(`${entry.sheetIndex}:${formula.sharedIndex}`, entry);
  }

  const formulas: ResolvedWorkbookFormula[] = [];
  const seen = new Set<string>();
  for (const entry of cells) {
    const formula = formulaOptions(entry.cell);
    if (!formula) continue;
    let source = formula.formula;
    if (formula.type === "shared" && formula.sharedIndex !== undefined) {
      const master = shared.get(`${entry.sheetIndex}:${formula.sharedIndex}`);
      if (!master) continue;
      const masterFormula = formulaOptions(master.cell)!;
      source = resolveSharedFormula(
        {
          formula: masterFormula.formula,
          masterRow: master.row,
          masterColumn: master.column,
        },
        { row: entry.row, column: entry.column },
        dialect,
      ).raw;
    }
    if (!source || seen.has(entry.reference)) continue;
    seen.add(entry.reference);
    const followers =
      formula.type === "array" && formula.reference
        ? cellsInReference(cells, entry.sheetIndex, formula.reference)
        : undefined;
    formulas.push({
      id: `Sheet${entry.sheetIndex}!${entry.reference}`,
      source: entry,
      formula: source,
      type: formula.type ?? "normal",
      ast: parseWorkbookFormula(source, dialect),
      followers,
    });
  }
  return formulas;
}

function distributeFormulaValue(
  targets: readonly WorkbookCell[],
  value: FormulaValue,
  type: Required<FormulaOptions>["type"],
): void {
  const matrix = isMatrix(value) ? value : type === "array" ? [[value]] : undefined;
  if (!matrix) {
    assignFormulaCache(targets[0]!.cell, firstScalar(value));
    return;
  }
  targets.forEach((target, index) => {
    const row = Math.floor(index / (matrix[0]?.length || 1));
    const column = index % (matrix[0]?.length || 1);
    assignFormulaCache(target.cell, matrix[row]?.[column] ?? null);
  });
}

function assignFormulaCache(cell: CellOptions, value: FormulaScalar): void {
  if (value instanceof FormulaError) {
    cell.error = value.value;
    return;
  }
  if (typeof value === "string" && isFormulaErrorToken(value)) {
    cell.error = value;
    return;
  }
  if (cell.error !== undefined) delete cell.error;
  if (value === null) return;
  cell.value = value;
}

function cellsInReference(
  cells: readonly WorkbookCell[],
  sheetIndex: number,
  reference: string,
): WorkbookCell[] {
  const range = parseRange(reference);
  if (!range) return [];
  const byPosition = new Map(
    cells
      .filter((entry) => entry.sheetIndex === sheetIndex)
      .map((entry) => [cellKey(0, entry.row, entry.column), entry]),
  );
  const found: WorkbookCell[] = [];
  for (let row = range.firstRow; row <= range.lastRow; row += 1)
    for (let column = range.firstColumn; column <= range.lastColumn; column += 1) {
      const foundCell = byPosition.get(cellKey(0, row, column));
      if (foundCell) found.push(foundCell);
    }
  return found;
}

function parseWorkbookFormula(
  formula: string,
  dialect: WorkbookFormulaDialect = "excel",
): ReturnType<typeof parseFormula> {
  return parseFormula(formula.startsWith("=") ? formula : `=${formula}`, { dialect });
}

function formulaOptions(cell: CellOptions): FormulaOptions | undefined {
  return typeof cell.formula === "string"
    ? { formula: cell.formula, type: "normal" }
    : cell.formula;
}

function toFormulaScalar(value: CellOptions["value"]): FormulaScalar {
  if (value === null || value === undefined) return null;
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean")
    return value;
  if (value instanceof Date) return value.toISOString();
  if ("runs" in value) return value.runs?.map((run) => run.text).join("") || value.text || "";
  return value.text ?? "";
}

function normalizeErrorToken(value: string): FormulaScalar {
  const token = value.replace(/^#+/u, ERROR_PREFIX).toUpperCase();
  return token === "N/A" ? "#N/A" : token;
}

function parseRange(reference: string):
  | {
      firstRow: number;
      firstColumn: number;
      lastRow: number;
      lastColumn: number;
    }
  | undefined {
  const [start, end] = reference.split(":");
  const from = parseReference(start ?? "");
  const to = end === undefined ? from : parseReference(end);
  if (!from || !to) return undefined;
  return {
    firstRow: Math.min(from.row, to.row),
    firstColumn: Math.min(from.column, to.column),
    lastRow: Math.max(from.row, to.row),
    lastColumn: Math.max(from.column, to.column),
  };
}

function parseReference(reference: string): { row: number; column: number } | undefined {
  const match = COLUMN_PATTERN.exec(reference.toUpperCase());
  if (!match) return undefined;
  let column = 0;
  for (const character of match[2]!) column = column * 26 + character.charCodeAt(0) - 64;
  const row = Number(match[4]);
  if (row < 1 || row > 1_048_576 || column < 1 || column > 16_384) return undefined;
  return { row, column };
}

function formatReference(row: number, column: number): string {
  let text = "";
  let value = column;
  while (value > 0) {
    text = String.fromCharCode(65 + ((value - 1) % 26)) + text;
    value = Math.floor((value - 1) / 26);
  }
  return `${text}${row}`;
}

function worksheetName(options: CanonicalWorkbookOptions, sheetIndex: number): string {
  return options.worksheets?.[sheetIndex]?.name ?? `Sheet${sheetIndex + 1}`;
}

function actualRowNumber(row: RowOptions): number | undefined {
  if (row.rowNumber !== undefined) return row.rowNumber;
  const reference = row.cells?.find((cell) => cell.reference)?.reference;
  return reference ? parseReference(reference)?.row : undefined;
}

function tableName(table: TableOptions): string {
  return (table.name ?? table.displayName).toLowerCase();
}

function cellKey(sheet: number, row: number, column: number): string {
  return `${sheet}:${row}:${column}`;
}

function firstScalar(value: FormulaValue): FormulaScalar {
  return isMatrix(value) ? (value[0]?.[0] ?? null) : value;
}

function isMatrix(value: FormulaValue): value is FormulaMatrix {
  return Array.isArray(value);
}
