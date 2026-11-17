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
  const tableXmlById = new Map<TableOptions, string>();
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
    const occupied = new Set<string>();
    let maximumColumns = explicitColumns.length;
    const rows = current.rows.map((row, rowIndex) => {
      const cells: string[] = [];
      let column = 0;
      const isOccupied = (columnIndex: number): boolean =>
        occupied.has(`${rowIndex}:${columnIndex}`);
      const addCovered = (columnIndex: number): void => {
        cells.push(xmlElement("table:covered-table-cell"));
        occupied.delete(`${rowIndex}:${columnIndex}`);
      };
      for (const cell of "cells" in row ? row.cells : []) {
        while (isOccupied(column)) {
          addCovered(column);
          column += 1;
        }
        const columnSpan = "columnSpan" in cell ? Math.max(1, cell.columnSpan ?? 1) : 1;
        const rowSpan = "rowSpan" in cell ? Math.max(1, cell.rowSpan ?? 1) : 1;
        const children =
          "children" in cell
            ? cell.children.map((child) =>
                "table" in child
                  ? (tableXmlById.get(child.table) ?? renderBlock(child))
                  : renderBlock(child),
              )
            : [];
        cells.push(
          xmlElement(
            "table:table-cell",
            {
              "office:value-type": "string",
              "table:number-columns-spanned": columnSpan > 1 ? columnSpan : undefined,
              "table:number-rows-spanned": rowSpan > 1 ? rowSpan : undefined,
            },
            children,
          ),
        );
        for (let offset = 1; offset < columnSpan; offset += 1) {
          cells.push(xmlElement("table:covered-table-cell"));
        }
        for (let rowOffset = 1; rowOffset < rowSpan; rowOffset += 1) {
          for (let columnOffset = 0; columnOffset < columnSpan; columnOffset += 1) {
            occupied.add(`${rowIndex + rowOffset}:${column + columnOffset}`);
          }
        }
        column += columnSpan;
        maximumColumns = Math.max(maximumColumns, column);
      }
      while (isOccupied(column)) {
        addCovered(column);
        column += 1;
      }
      maximumColumns = Math.max(maximumColumns, column);
      return xmlElement("table:table-row", undefined, [cells.join("")]);
    });
    const columns = Math.max(maximumColumns, 1);
    const columnElements = Array.from(
      { length: Math.max(columns, 1) },
      (_, index) => explicitColumns[index] ?? xmlElement("table:table-column"),
    );
    const bodyRows = rows.length
      ? rows.join("")
      : xmlElement("table:table-row", undefined, [xmlElement("table:table-cell")]);
    tableXmlById.set(
      current,
      xmlElement(
        "table:table",
        {
          "table:name": tableName.get(current)!,
          "table:style-name": current.style,
        },
        [...columnElements, bodyRows],
      ),
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
          ...(attributeNumber(cell, "table:number-columns-spanned") === undefined
            ? {}
            : { columnSpan: attributeNumber(cell, "table:number-columns-spanned") }),
          ...(attributeNumber(cell, "table:number-rows-spanned") === undefined
            ? {}
            : { rowSpan: attributeNumber(cell, "table:number-rows-spanned") }),
        };
      }),
    }));
    parsedTables.set(current, {
      table: {
        columnWidths,
        rows,
        ...(attributeString(current, "table:style-name") === undefined
          ? {}
          : { style: attributeString(current, "table:style-name") }),
      },
    });
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
