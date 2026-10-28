import {
  type FormulaAst,
  type FormulaRangeEdge,
  type FormulaReferenceInfo,
  type FormulaTableSelector,
  type FormulaWorkbookIdentity,
} from "./ast";
import { parseFormula } from "./parser";
import {
  isWholeColumn,
  isWholeRow,
  maximumColumns,
  maximumRows,
  numberToColumn,
} from "./references";

export interface FormulaDependencyTarget {
  kind: "cell" | "range" | "name" | "table" | "unknown";
  workbook?: FormulaWorkbookIdentity;
  sheet?: string;
  toSheet?: string;
  from?: { row: number; column: number };
  to?: { row: number; column: number };
  spill?: boolean;
  name?: string;
  table?: FormulaTableSelector;
  implicitIntersection?: boolean;
}

export interface FormulaDependency {
  target: FormulaDependencyTarget;
  sourceStart: number;
  sourceEnd: number;
}

export function formatFormulaDependencyTarget(target: FormulaDependencyTarget): string {
  if (target.kind === "name") return target.name ?? "#NAME?";
  if (target.kind === "table") return target.name ?? "#REF!";
  if (!target.from) return target.sheet ? `${formatSheet(target.sheet)}!#REF!` : "#REF!";
  const start = `${numberToColumn(target.from.column)}${target.from.row}`;
  const end = target.to ? `${numberToColumn(target.to.column)}${target.to.row}` : undefined;
  const range = `${end && end !== start ? `${start}:${end}` : start}${target.spill ? "#" : ""}`;
  const sheets =
    target.sheet === undefined
      ? ""
      : target.toSheet === undefined
        ? `${formatSheet(target.sheet)}!`
        : `${formatSheet(target.sheet)}:${formatSheet(target.toSheet)}!`;
  return `${formatWorkbookPrefix(target.workbook)}${sheets}${range}`;
}

function formatWorkbookPrefix(workbook: FormulaWorkbookIdentity | undefined): string {
  if (workbook === undefined) return "";
  if (typeof workbook === "string") return `[${workbook}]`;
  if (workbook.path !== undefined) return `[${workbook.path}]`;
  if (workbook.index !== undefined) return `[${workbook.index}]`;
  return workbook.name === undefined ? "" : `[${workbook.name}]`;
}

function formatSheet(name: string): string {
  return /[^A-Za-z0-9_]/u.test(name) ? `'${name.replaceAll("'", "''")}'` : name;
}

function referenceTargets(reference: FormulaReferenceInfo): FormulaDependencyTarget[] {
  if (reference.table)
    return [
      {
        kind: "table",
        name: reference.table.table,
        table: reference.table,
        ...(reference.implicitIntersection ? { implicitIntersection: true } : {}),
      },
    ];
  if (reference.spill) {
    const area = reference.areas?.[0];
    const anchor = area?.start;
    if (!anchor) return [{ kind: "unknown" }];
    return [
      {
        kind: "cell",
        ...(reference.workbook ? { workbook: reference.workbook } : {}),
        ...(area?.fromSheet ? { sheet: area.fromSheet } : {}),
        from: edgeTarget(anchor, "from"),
        spill: true,
      },
    ];
  }
  if (reference.areas) {
    return reference.areas.flatMap((area) => {
      const from = edgeTarget(area.start, "from");
      const to = edgeTarget(area.end ?? area.start, "to");
      const workbook = area.workbook ?? reference.workbook;
      const sheet =
        area.fromSheet ??
        (reference.sheets?.names.length === 1 ? reference.sheets.names[0] : undefined);
      return [
        {
          kind: from && to && (from.row !== to.row || from.column !== to.column) ? "range" : "cell",
          ...(workbook ? { workbook } : {}),
          ...(sheet ? { sheet } : {}),
          ...(area.toSheet ? { toSheet: area.toSheet } : {}),
          ...(from ? { from } : {}),
          ...(to ? { to } : {}),
        },
      ];
    });
  }
  const sheet = reference.sheets?.names[0];
  if (reference.start) {
    const from = { row: reference.start.row, column: reference.start.column };
    const sheetTarget = sheet === undefined ? {} : { sheet };
    if (reference.end) {
      return [
        {
          kind: "range",
          ...(reference.workbook ? { workbook: reference.workbook } : {}),
          ...sheetTarget,
          from,
          to: { row: reference.end.row, column: reference.end.column },
        },
      ];
    }
    return [
      {
        kind: "cell",
        ...(reference.workbook ? { workbook: reference.workbook } : {}),
        ...sheetTarget,
        from,
      },
    ];
  }
  return [{ kind: sheet ? "range" : "unknown", ...(sheet === undefined ? {} : { sheet }) }];
}

function edgeTarget(
  edge: FormulaRangeEdge | undefined,
  boundary: "from" | "to",
): { row: number; column: number } | undefined {
  if (!edge) return undefined;
  if (isWholeColumn(edge))
    return { row: boundary === "from" ? 1 : maximumRows, column: edge.column };
  if (isWholeRow(edge)) return { row: edge.row, column: boundary === "from" ? 1 : maximumColumns };
  return { row: edge.row, column: edge.column };
}

function staticOffsetValue(argument: FormulaAst | undefined, fallback: number): number | undefined {
  if (argument === undefined || argument.kind === "missing") return fallback;
  if (argument.kind === "number" && Number.isInteger(argument.value)) return argument.value;
  return undefined;
}

function indirectUsesA1(style: FormulaAst | undefined): boolean | undefined {
  if (style === undefined) return true;
  if (style.kind === "boolean") return style.value;
  if (style.kind === "number") return style.value !== 0;
  if (style.kind !== "string") return undefined;
  const value = style.value.toUpperCase();
  if (value === "TRUE") return true;
  if (value === "FALSE") return false;
  return undefined;
}

function offsetDependencies(ast: Extract<FormulaAst, { kind: "call" }>): FormulaDependency[] {
  const base = ast.args[0];
  const rows = ast.args[1];
  const columns = ast.args[2];
  if (base?.kind !== "reference")
    return [{ target: { kind: "unknown" }, sourceStart: ast.start, sourceEnd: ast.end }];
  const anchor = base.reference?.areas?.[0];
  const start = anchor?.start;
  const end = anchor?.end ?? start;
  if (
    !start ||
    !end ||
    isWholeColumn(start) ||
    isWholeRow(start) ||
    isWholeColumn(end) ||
    isWholeRow(end)
  )
    return [{ target: { kind: "unknown" }, sourceStart: ast.start, sourceEnd: ast.end }];
  const rowsValue = staticOffsetValue(rows, 0);
  const columnsValue = staticOffsetValue(columns, 0);
  const height = staticOffsetValue(ast.args[3], Math.abs(end.row - start.row) + 1);
  const width = staticOffsetValue(ast.args[4], Math.abs(end.column - start.column) + 1);
  if (
    rowsValue === undefined ||
    columnsValue === undefined ||
    height === undefined ||
    width === undefined ||
    height <= 0 ||
    width <= 0
  )
    return [{ target: { kind: "unknown" }, sourceStart: ast.start, sourceEnd: ast.end }];
  const firstRow = start.row + rowsValue;
  const firstColumn = start.column + columnsValue;
  const target: FormulaDependencyTarget = {
    kind: height === 1 && width === 1 ? "cell" : "range",
    workbook: anchor?.workbook ?? base.reference?.workbook,
    ...(anchor?.fromSheet ? { sheet: anchor.fromSheet } : {}),
    from: { row: firstRow, column: firstColumn },
    to: { row: firstRow + height - 1, column: firstColumn + width - 1 },
  };
  return [{ target, sourceStart: base.start, sourceEnd: ast.end }];
}

export function collectDependencies(ast: FormulaAst): FormulaDependency[] {
  const dependencies: FormulaDependency[] = [];
  collectScopedDependencies(ast, new Set(), dependencies);
  return dependencies;
}

function indirectDependencies(ast: Extract<FormulaAst, { kind: "call" }>): FormulaDependency[] {
  const unknown: FormulaDependency = {
    target: { kind: "unknown" },
    sourceStart: ast.start,
    sourceEnd: ast.end,
  };
  const source = ast.args[0];
  const style = ast.args[1];
  const usesA1 = indirectUsesA1(style);
  if (usesA1 === false || source?.kind !== "string") return [unknown];
  try {
    const target = parseFormula(source.value);
    if (target.kind !== "reference" && target.kind !== "name") return [unknown];
    const dependencies: FormulaDependency[] = [];
    collectScopedDependencies(target, new Set(), dependencies);
    return dependencies;
  } catch {
    return [unknown];
  }
}

function collectScopedDependencies(
  ast: FormulaAst,
  locals: ReadonlySet<string>,
  dependencies: FormulaDependency[],
): void {
  if (ast.kind === "reference") {
    if (!locals.has(ast.raw.toUpperCase()) && ast.reference) {
      for (const target of referenceTargets(ast.reference))
        dependencies.push({ target, sourceStart: ast.start, sourceEnd: ast.end });
    }
    return;
  }
  if (ast.kind === "name") {
    if (locals.has(ast.raw.toUpperCase())) return;
    dependencies.push({
      target: {
        kind: "name",
        name: ast.name,
        ...(ast.implicitIntersection ? { implicitIntersection: true } : {}),
      },
      sourceStart: ast.start,
      sourceEnd: ast.end,
    });
    return;
  }
  if (ast.kind === "array") {
    for (const row of ast.rows)
      for (const child of row) collectScopedDependencies(child, locals, dependencies);
    return;
  }
  if (ast.kind === "group" || ast.kind === "unary" || ast.kind === "postfix") {
    collectScopedDependencies(ast.operand, locals, dependencies);
    return;
  }
  if (ast.kind === "binary") {
    collectScopedDependencies(ast.left, locals, dependencies);
    collectScopedDependencies(ast.right, locals, dependencies);
    return;
  }
  if (ast.kind === "application") {
    collectScopedDependencies(ast.callee, locals, dependencies);
    for (const child of ast.args) collectScopedDependencies(child, locals, dependencies);
    return;
  }
  if (ast.kind !== "call") return;

  const name = ast.name.toUpperCase();
  if (name === "LAMBDA" && ast.args.length >= 2) {
    const parameters = ast.args.slice(0, -1).map((argument) => argument.raw.toUpperCase());
    const scopedLocals = new Set([...locals, ...parameters]);
    collectScopedDependencies(ast.args.at(-1)!, scopedLocals, dependencies);
    return;
  }
  if (name === "LET" && ast.args.length >= 3 && ast.args.length % 2 === 1) {
    let scopedLocals = locals;
    for (let index = 0; index + 1 < ast.args.length; index += 2) {
      collectScopedDependencies(ast.args[index + 1]!, scopedLocals, dependencies);
      scopedLocals = new Set([...scopedLocals, ast.args[index]!.raw.toUpperCase()]);
    }
    collectScopedDependencies(ast.args.at(-1)!, scopedLocals, dependencies);
    return;
  }
  if (name === "OFFSET") {
    dependencies.push(...offsetDependencies(ast));
    for (const child of ast.args.slice(1))
      if (staticOffsetValue(child, 0) === undefined)
        collectScopedDependencies(child, locals, dependencies);
    return;
  }
  if (name === "INDIRECT") {
    dependencies.push(...indirectDependencies(ast));
    return;
  }
  for (const child of ast.args) collectScopedDependencies(child, locals, dependencies);
}

export interface FormulaDependencyNode {
  id: string;
  ast: FormulaAst;
  /** Defined-name node name; cell nodes omit it. */
  name?: string;
  /** Sheet that scopes a defined-name node; omitted for workbook scope. */
  scopeSheet?: string;
  cell?: {
    workbook?: FormulaWorkbookIdentity;
    sheet?: string;
    row: number;
    column: number;
  };
}

export interface FormulaDependencyGraph {
  nodes: Map<string, FormulaDependencyNode>;
  edges: Map<string, Set<string>>;
}

export interface FormulaDependencyGraphOptions {
  sheetOrder?: readonly string[];
  defaultWorkbook?: FormulaWorkbookIdentity;
  defaultSheet?: string;
  getSheetId?(workbook: FormulaWorkbookIdentity | undefined, sheet: string): number | undefined;
  getTableName?(name: string):
    | {
        workbook?: FormulaWorkbookIdentity;
        sheet?: string;
        range: { firstRow: number; firstColumn: number; lastRow: number; lastColumn: number };
        columnNames?: string[];
        hasHeaders?: boolean;
        hasTotals?: boolean;
      }
    | undefined;
  getSpillRange?(
    workbook: FormulaWorkbookIdentity | undefined,
    sheet: string | undefined,
    row: number,
    column: number,
  ):
    | {
        firstRow: number;
        firstColumn: number;
        lastRow: number;
        lastColumn: number;
      }
    | undefined;
}

interface FormulaTableGraphInfo {
  workbook?: FormulaWorkbookIdentity;
  sheet?: string;
  range: { firstRow: number; firstColumn: number; lastRow: number; lastColumn: number };
  columnNames?: string[];
  hasHeaders?: boolean;
  hasTotals?: boolean;
}

function sameWorkbook(
  left: FormulaWorkbookIdentity | undefined,
  right: FormulaWorkbookIdentity | undefined,
): boolean {
  const leftKey = workbookKey(left);
  const rightKey = workbookKey(right);
  return leftKey === rightKey;
}

function workbookKey(workbook: FormulaWorkbookIdentity | undefined): string {
  if (workbook === undefined) return "";
  if (typeof workbook === "string") return `name\0${workbook}`;
  if (workbook.path === undefined && workbook.index === undefined)
    return `name\0${workbook.name ?? ""}`;
  return `source\0${workbook.index ?? ""}\0${workbook.path ?? ""}\0${workbook.name ?? ""}`;
}

function tableDependency(
  target: FormulaDependencyTarget,
  table: FormulaTableGraphInfo,
  caller: { workbook?: FormulaWorkbookIdentity; sheet?: string; row?: number },
): FormulaDependencyTarget[] {
  const selector = target.table;
  const resolveColumn = (column: string): number => {
    const index = table.columnNames?.findIndex(
      (name) => name.toLowerCase() === column.toLowerCase(),
    );
    return index === undefined || index < 0 ? 0 : table.range.firstColumn + index;
  };
  const selected: { from: number; to: number }[] = [];
  for (const part of selector?.selectors ?? []) {
    const indexes =
      part.kind === "column"
        ? [resolveColumn(part.column)]
        : [resolveColumn(part.start), resolveColumn(part.end)];
    if (indexes.some((index) => index === 0)) return [{ kind: "unknown" }];
    selected.push({ from: Math.min(...indexes), to: Math.max(...indexes) });
  }
  const item = target.table?.item ?? (target.implicitIntersection ? "this-row" : undefined);
  const bodyFirst = table.range.firstRow + (table.hasHeaders ? 1 : 0);
  const bodyLast = table.range.lastRow - (table.hasTotals ? 1 : 0);
  let firstRow: number;
  let lastRow: number;
  if (item === "headers") {
    if (!table.hasHeaders) return [];
    firstRow = table.range.firstRow;
    lastRow = firstRow;
  } else if (item === "totals") {
    if (!table.hasTotals) return [];
    firstRow = table.range.lastRow;
    lastRow = firstRow;
  } else if (item === "this-row") {
    if (
      caller.row === undefined ||
      !sameWorkbook(caller.workbook, table.workbook) ||
      caller.sheet !== table.sheet ||
      caller.row < bodyFirst ||
      caller.row > bodyLast
    )
      return [];
    firstRow = caller.row;
    lastRow = firstRow;
  } else if (item === "all") {
    firstRow = table.range.firstRow;
    lastRow = table.range.lastRow;
  } else {
    firstRow = bodyFirst;
    lastRow = bodyLast;
    if (firstRow > lastRow) return [];
  }
  return (
    selected.length ? selected : [{ from: table.range.firstColumn, to: table.range.lastColumn }]
  ).map(({ from, to }): FormulaDependencyTarget => ({
    kind: "range",
    ...(table.workbook ? { workbook: table.workbook } : {}),
    ...(table.sheet ? { sheet: table.sheet } : {}),
    from: { row: firstRow, column: from },
    to: { row: lastRow, column: to },
  }));
}

export function buildDependencyGraph(
  nodes: FormulaDependencyNode[],
  options: FormulaDependencyGraphOptions = {},
): FormulaDependencyGraph {
  const graph: FormulaDependencyGraph = {
    nodes: new Map(nodes.map((node) => [node.id, node])),
    edges: new Map(nodes.map((node) => [node.id, new Set<string>()])),
  };
  const graphOptions = containsMultipleWorkbooks(nodes, options)
    ? { ...options, sheetOrder: undefined }
    : options;
  const nodeIdsByName = new Map(nodes.map((node) => [node.id.toUpperCase(), node.id]));
  const namedNodes = nodes.filter((node) => node.name !== undefined);
  const locatedNodes = nodes.filter((node) => node.cell !== undefined);
  for (const node of nodes) {
    for (const dependency of collectDependencies(node.ast)) {
      const name = dependency.target.kind === "name" ? dependency.target.name : undefined;
      const namedNode = name
        ? namedNodes.find(
            (node) =>
              node.name?.toUpperCase() === name.toUpperCase() &&
              (node.scopeSheet === undefined ||
                node.scopeSheet === (dependency.target.sheet ?? node.cell?.sheet)),
          )
        : undefined;
      const nodeId = name ? (namedNode?.id ?? nodeIdsByName.get(name.toUpperCase())) : undefined;
      if (name && nodeId !== undefined) {
        graph.edges.get(node.id)?.add(nodeId);
        continue;
      }
      let target = dependency.target;
      if (name && nodeId === undefined && options.getTableName?.(name)) {
        target = {
          kind: "table",
          name,
          table: {
            table: name,
            selectors: [],
            ...(dependency.target.implicitIntersection ? { item: "this-row" } : {}),
          },
        };
      }
      if (target.kind === "table") {
        const table = options.getTableName?.(target.name ?? "");
        if (!table) {
          for (const other of nodes)
            if (other.id !== node.id) graph.edges.get(node.id)?.add(other.id);
          continue;
        }
        const normalizedTable: FormulaTableGraphInfo = {
          ...table,
          workbook: table.workbook ?? node.cell?.workbook ?? options.defaultWorkbook,
          sheet: table.sheet ?? node.cell?.sheet ?? options.defaultSheet,
        };
        const resolvedTargets = tableDependency(target, normalizedTable, {
          workbook: node.cell?.workbook ?? options.defaultWorkbook,
          sheet: node.cell?.sheet ?? options.defaultSheet,
          row: node.cell?.row,
        });
        for (const resolved of resolvedTargets) {
          resolved.sheet = resolved.sheet ?? node.cell?.sheet ?? options.defaultSheet;
          if (!resolved.from) continue;
          connectRange(
            graph,
            node,
            resolved,
            locatedNodes,
            graphOptions,
            resolved.from,
            resolved.to ?? resolved.from,
          );
        }
        continue;
      }
      if (target.kind === "unknown") {
        for (const other of nodes)
          if (other.id !== node.id) graph.edges.get(node.id)?.add(other.id);
        continue;
      }
      if ((target.kind !== "cell" && target.kind !== "range") || !target.from) continue;
      const workbook = target.workbook ?? node.cell?.workbook ?? options.defaultWorkbook;
      const sheet = target.sheet ?? node.cell?.sheet ?? options.defaultSheet;
      const expanded: FormulaDependencyTarget = { ...target, workbook, sheet };
      if (expanded.spill) {
        const anchor = expanded.from;
        if (!anchor) {
          for (const other of nodes)
            if (other.id !== node.id) graph.edges.get(node.id)?.add(other.id);
          continue;
        }
        const spill = options.getSpillRange?.(workbook, sheet, anchor.row, anchor.column);
        if (spill) {
          expanded.from = { row: spill.firstRow, column: spill.firstColumn };
          expanded.to = { row: spill.lastRow, column: spill.lastColumn };
        } else {
          for (const other of nodes)
            if (other.id !== node.id) graph.edges.get(node.id)?.add(other.id);
          continue;
        }
      }
      const from = expanded.from;
      if (!from) continue;
      const to = expanded.to ?? from;
      connectRange(graph, node, expanded, locatedNodes, graphOptions, from, to);
    }
  }
  return graph;
}

function containsMultipleWorkbooks(
  nodes: readonly FormulaDependencyNode[],
  options: FormulaDependencyGraphOptions,
): boolean {
  const workbooks = new Set<string>();
  if (options.defaultWorkbook !== undefined) workbooks.add(workbookKey(options.defaultWorkbook));
  for (const node of nodes) {
    if (node.cell?.workbook !== undefined) workbooks.add(workbookKey(node.cell.workbook));
    for (const dependency of collectDependencies(node.ast))
      if (dependency.target.workbook !== undefined)
        workbooks.add(workbookKey(dependency.target.workbook));
  }
  return workbooks.size > 1;
}

function connectRange(
  graph: FormulaDependencyGraph,
  source: FormulaDependencyNode,
  target: FormulaDependencyTarget,
  locatedNodes: readonly FormulaDependencyNode[],
  options: FormulaDependencyGraphOptions,
  from: { row: number; column: number },
  to: { row: number; column: number },
): void {
  for (const located of locatedNodes) {
    if (!located.cell) continue;
    const workbook = located.cell.workbook ?? options.defaultWorkbook;
    if (!sheetMatches(target, located.cell.sheet, options, workbook)) continue;
    if (
      located.cell.row < from.row ||
      located.cell.row > to.row ||
      located.cell.column < from.column ||
      located.cell.column > to.column
    )
      continue;
    graph.edges.get(source.id)?.add(located.id);
  }
}

function sheetMatches(
  target: FormulaDependencyTarget,
  sheet: string | undefined,
  options: FormulaDependencyGraphOptions,
  workbook?: FormulaWorkbookIdentity,
): boolean {
  if (!sameWorkbook(target.workbook, workbook)) return false;
  if (target.toSheet === undefined) return target.sheet === undefined || target.sheet === sheet;
  if (target.sheet === undefined || sheet === undefined) return false;
  if (options.sheetOrder) {
    const sheetOrder = options.sheetOrder;
    const fromIndex = sheetOrder.indexOf(target.sheet);
    const toIndex = sheetOrder.indexOf(target.toSheet);
    const sheetIndex = sheetOrder.indexOf(sheet);
    if (fromIndex < 0 || toIndex < 0 || sheetIndex < 0) return true;
    return sheetIndex >= Math.min(fromIndex, toIndex) && sheetIndex <= Math.max(fromIndex, toIndex);
  }
  if (options.getSheetId) {
    const fromId = options.getSheetId(target.workbook, target.sheet);
    const toId = options.getSheetId(target.workbook, target.toSheet);
    const sheetId = options.getSheetId(target.workbook, sheet);
    if (fromId === undefined || toId === undefined || sheetId === undefined) return true;
    return sheetId >= Math.min(fromId, toId) && sheetId <= Math.max(fromId, toId);
  }
  return true;
}

export function detectFormulaCycles(graph: FormulaDependencyGraph): string[][] {
  const indices = new Map<string, number>();
  const lowLinks = new Map<string, number>();
  const stack: string[] = [];
  const onStack = new Set<string>();
  const cycles: string[][] = [];
  let nextIndex = 0;

  function start(id: string): void {
    indices.set(id, nextIndex);
    lowLinks.set(id, nextIndex);
    nextIndex += 1;
    stack.push(id);
    onStack.add(id);
  }

  function extractComponent(id: string): string[] {
    const component: string[] = [];
    for (;;) {
      const member = stack.pop()!;
      onStack.delete(member);
      component.push(member);
      if (member === id) break;
    }
    return component.sort();
  }

  for (const root of graph.nodes.keys()) {
    if (indices.has(root)) continue;
    start(root);
    const frames: { id: string; dependencies: string[]; cursor: number }[] = [
      { id: root, dependencies: [...(graph.edges.get(root) ?? [])], cursor: 0 },
    ];
    while (frames.length > 0) {
      const frame = frames.at(-1)!;
      if (frame.cursor < frame.dependencies.length) {
        const dependency = frame.dependencies[frame.cursor]!;
        frame.cursor += 1;
        if (!indices.has(dependency)) {
          start(dependency);
          frames.push({
            id: dependency,
            dependencies: [...(graph.edges.get(dependency) ?? [])],
            cursor: 0,
          });
        } else if (onStack.has(dependency)) {
          lowLinks.set(frame.id, Math.min(lowLinks.get(frame.id)!, indices.get(dependency)!));
        }
        continue;
      }
      frames.pop();
      if (lowLinks.get(frame.id) === indices.get(frame.id)) {
        const component = extractComponent(frame.id);
        if (component.length > 1 || graph.edges.get(frame.id)?.has(frame.id))
          cycles.push(component);
      }
      const parent = frames.at(-1);
      if (parent)
        lowLinks.set(parent.id, Math.min(lowLinks.get(parent.id)!, lowLinks.get(frame.id)!));
    }
  }
  return cycles.sort((left, right) => left.join("\u0000").localeCompare(right.join("\u0000")));
}

export function topologicalFormulaOrder(graph: FormulaDependencyGraph): string[] {
  if (detectFormulaCycles(graph).length > 0) throw new Error("cyclic formula dependency");
  const order: string[] = [];
  const visited = new Set<string>();
  const pending: Array<{ id: string; dependencies: Iterator<string> }> = [];

  for (const id of graph.nodes.keys()) {
    if (visited.has(id)) continue;
    pending.push({ id, dependencies: (graph.edges.get(id) ?? new Set<string>()).values() });
    while (pending.length > 0) {
      const frame = pending.at(-1)!;
      const step = frame.dependencies.next();
      if (step.done) {
        pending.pop();
        visited.add(frame.id);
        order.push(frame.id);
        continue;
      }
      const dependency = step.value;
      if (visited.has(dependency)) continue;
      pending.push({
        id: dependency,
        dependencies: (graph.edges.get(dependency) ?? new Set<string>()).values(),
      });
    }
  }
  return order;
}

export function dependentsByFormula(graph: FormulaDependencyGraph): Map<string, Set<string>> {
  const dependents = new Map([...graph.nodes.keys()].map((id) => [id, new Set<string>()]));
  for (const [formula, dependencies] of graph.edges)
    for (const dependency of dependencies) dependents.get(dependency)?.add(formula);
  return dependents;
}

export function collectDirtyFormulas(
  graph: FormulaDependencyGraph,
  changed: readonly string[],
): Set<string> {
  const dependents = dependentsByFormula(graph);
  const dirty = new Set<string>();
  const queue = [...changed];
  while (queue.length > 0) {
    const id = queue.pop();
    if (id === undefined || dirty.has(id)) continue;
    dirty.add(id);
    queue.push(...(dependents.get(id) ?? []));
  }
  return dirty;
}
