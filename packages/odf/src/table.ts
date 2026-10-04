import type { SectionChild, TableOptions } from "@office-open/docx";
import type { Element } from "@office-open/xml";

import { attributeNumber, attributeString, childrenNamed, lengthToEmu, xmlElement } from "./xml";

/** Serializes a docx table to table:table XML shared by ODT and ODP bodies. */
export function tableXml(
  table: TableOptions,
  styles: string[],
  renderBlock: (child: SectionChild) => string,
): string {
  const explicitColumns = (table.columnWidths ?? []).map((width) => {
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
  const rows = table.rows.map((row) => {
    const cells = ("cells" in row ? row.cells : []).map((cell) => {
      const span = "columnSpan" in cell ? cell.columnSpan : undefined;
      const children = "children" in cell ? cell.children.map((child) => renderBlock(child)) : [];
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
    });
    return xmlElement("table:table-row", undefined, cells);
  });
  const rowSpans = table.rows.map((row) =>
    ("cells" in row ? row.cells : []).reduce(
      (total, cell) => total + ("columnSpan" in cell ? (cell.columnSpan ?? 1) : 1),
      0,
    ),
  );
  const columnCount = Math.max(explicitColumns.length, ...rowSpans, 1);
  const columns = Array.from(
    { length: columnCount },
    (_, column) => explicitColumns[column] ?? xmlElement("table:table-column"),
  );
  const bodyRows = rows.length
    ? rows
    : [xmlElement("table:table-row", undefined, [xmlElement("table:table-cell")])];
  return xmlElement("table:table", { "table:name": `Table${styles.length + 1}` }, [
    columns.join(""),
    bodyRows.join(""),
  ]);
}

/** Parses a table:table element back to a docx table block. */
export function parseTable(
  element: Element,
  columnWidthOf: (styleName: string) => number | undefined,
  parseBlock: (element: Element) => SectionChild,
): SectionChild {
  const columnWidths = childrenNamed(element, "table:table-column").map(
    (column) => columnWidthOf(attributeString(column, "table:style-name") ?? "") ?? 5000,
  );
  const rows = childrenNamed(element, "table:table-row").map((row) => ({
    cells: childrenNamed(row, "table:table-cell").map((cell) => {
      const cellChildren =
        cell.elements
          ?.filter(
            (child) =>
              child.name === "text:p" || child.name === "text:h" || child.name === "table:table",
          )
          .map((child) => parseBlock(child)) ?? [];
      return {
        children: cellChildren,
        columnSpan: attributeNumber(cell, "table:number-columns-spanned"),
      };
    }),
  }));
  return { table: { columnWidths, rows } };
}
