import type { SectionChild, TableOptions } from "@office-open/docx";
import {
  attributeNumber,
  attributeString,
  childrenNamed,
  lengthToEmu,
  xmlElement,
} from "@office-open/odf";
import type { Element } from "@office-open/xml";

/** Serializes a docx table to table:table XML shared by ODT and ODP bodies. */
export function tableXml(
  table: TableOptions,
  styles: string[],
  renderBlock: (child: SectionChild) => string,
): string {
  const tables = collectTableTree(table);
  const columnXml = new Map<TableOptions, string[]>();
  const tableName = new Map<TableOptions, string>();
  const rowSpanCount = new Map<TableOptions, number[]>();
  for (const current of tables) {
    const explicitColumns = (current.columnWidths ?? []).map((width) => {
      const twips = typeof width === "number" ? width : Math.round(lengthToEmu(width)! / 635);
      const name = `T${styles.length + 1}`;
      styles.push(
        xmlElement("style:style", { "style:name": name, "style:family": "table-column" }, [
          xmlElement("style:table-column-properties", {
            "style:column-width": `${Number((twips / 567).toFixed(4))}cm`,
          }),
        ]),
      );
      return xmlElement("table:table-column", { "table:style-name": name });
    });
    columnXml.set(current, explicitColumns);
    tableName.set(current, `Table${styles.length + 1}`);
    rowSpanCount.set(
      current,
      current.rows.map((row) =>
        ("cells" in row ? row.cells : []).reduce(
          (total, cell) => total + ("columnSpan" in cell ? (cell.columnSpan ?? 1) : 1),
          0,
        ),
      ),
    );
  }
  const tableXmlById = new Map<TableOptions, string>();
  for (const current of tables) {
    const rows = current.rows.map((row) => {
      const cells = ("cells" in row ? row.cells : [])
        .map((cell) => {
          const span = "columnSpan" in cell ? cell.columnSpan : undefined;
          const children =
            "children" in cell
              ? cell.children.map((child) =>
                  "table" in child
                    ? (tableXmlById.get(child.table) ?? renderBlock(child))
                    : renderBlock(child),
                )
              : [];
          const xml = xmlElement(
            "table:table-cell",
            {
              "office:value-type": "string",
              "table:number-columns-spanned": span,
            },
            children,
          );
          const covered = Array.from({ length: Math.max(0, (span ?? 1) - 1) }, () =>
            xmlElement("table:covered-table-cell"),
          );
          return xml + covered.join("");
        })
        .join("");
      return xmlElement("table:table-row", undefined, [cells]);
    });
    const spans = rowSpanCount.get(current)!;
    const explicitColumns = columnXml.get(current)!;
    const columnCount = Math.max(explicitColumns.length, ...spans, 1);
    const columns = Array.from(
      { length: columnCount },
      (_, index) => explicitColumns[index] ?? xmlElement("table:table-column"),
    );
    const bodyRows = rows.length
      ? rows.join("")
      : xmlElement("table:table-row", undefined, [xmlElement("table:table-cell")]);
    tableXmlById.set(
      current,
      xmlElement("table:table", { "table:name": tableName.get(current)! }, [...columns, bodyRows]),
    );
  }
  return tableXmlById.get(table)!;
}

function collectTableTree(root: TableOptions): TableOptions[] {
  const order: TableOptions[] = [];
  const visited = new Set<TableOptions>();
  const pending: { readonly table: TableOptions; readonly expand: boolean }[] = [
    { table: root, expand: false },
  ];
  const expanded = new Set<TableOptions>();
  while (pending.length > 0) {
    const frame = pending.pop()!;
    if (frame.expand) {
      if (expanded.has(frame.table)) continue;
      expanded.add(frame.table);
      order.push(frame.table);
      continue;
    }
    if (visited.has(frame.table)) continue;
    visited.add(frame.table);
    pending.push({ table: frame.table, expand: true });
    for (const row of frame.table.rows) {
      if (!("cells" in row)) continue;
      for (const cell of row.cells) {
        if (!("children" in cell)) continue;
        for (const child of [...cell.children].reverse())
          if ("table" in child) pending.push({ table: child.table, expand: false });
      }
    }
  }
  return order;
}

/** Parses a table:table element back to a docx table block. */
export function parseTable(
  element: Element,
  columnWidthOf: (styleName: string) => number | undefined,
  parseBlock: (element: Element) => SectionChild,
): SectionChild {
  const parsedTables = new Map<Element, SectionChild>();
  for (const current of collectTableElements(element)) {
    const columnWidths = childrenNamed(current, "table:table-column").map(
      (column) => columnWidthOf(attributeString(column, "table:style-name") ?? "") ?? 5000,
    );
    const rows = childrenNamed(current, "table:table-row").map((row) => ({
      cells: childrenNamed(row, "table:table-cell").map((cell) => {
        const cellChildren =
          cell.elements
            ?.filter(
              (child) =>
                child.name === "text:p" || child.name === "text:h" || child.name === "table:table",
            )
            .map((child) =>
              child.name === "table:table"
                ? (parsedTables.get(child) ?? parseBlock(child))
                : parseBlock(child),
            ) ?? [];
        return {
          children: cellChildren,
          columnSpan: attributeNumber(cell, "table:number-columns-spanned"),
        };
      }),
    }));
    parsedTables.set(current, { table: { columnWidths, rows } });
  }
  return parsedTables.get(element)!;
}

function collectTableElements(root: Element): Element[] {
  const order: Element[] = [];
  const pending: { readonly element: Element; readonly expand: boolean }[] = [
    { element: root, expand: false },
  ];
  const expanded = new Set<Element>();
  while (pending.length > 0) {
    const frame = pending.pop()!;
    if (frame.expand) {
      if (expanded.has(frame.element)) continue;
      expanded.add(frame.element);
      order.push(frame.element);
      continue;
    }
    pending.push({ element: frame.element, expand: true });
    for (const row of [...childrenNamed(frame.element, "table:table-row")].reverse())
      for (const cell of [...childrenNamed(row, "table:table-cell")].reverse())
        for (const child of [...(cell.elements ?? [])].reverse())
          if (child.name === "table:table") pending.push({ element: child, expand: false });
  }
  return order;
}
