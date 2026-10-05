import type { SectionChild, TableOptions, TableCellOptions } from "@office-open/docx";

import { assertAllowed, reject, type RtfGenerateContext } from "./context";
import { control, rtfText } from "./escape";
import type { NumberingReferences } from "./numbering";
import { writeParagraph } from "./paragraph";

export function writeTableCells(
  cells: TableCellOptions[],
  context: RtfGenerateContext,
  references: NumberingReferences,
  path: string,
): string {
  return cells
    .map((cell, index) => {
      assertAllowed(cell, ["children"], "table", `${path}.cells[${index}]`);
      return writeCellChildren(
        cell.children,
        context,
        references,
        `${path}.cells[${index}].children`,
      );
    })
    .join("");
}

function writeCellChildren(
  children: SectionChild[],
  context: RtfGenerateContext,
  references: NumberingReferences,
  path: string,
): string {
  return children
    .map((child, index) => {
      if ("paragraph" in child) {
        if (typeof child.paragraph === "string") {
          assertAllowed(child, ["paragraph"], "table", `${path}[${index}]`);
          return `${rtfText(child.paragraph)}${control("par")}`;
        }
        return writeParagraph(child.paragraph, context, references, `${path}[${index}].paragraph`);
      }
      if ("table" in child) {
        return writeTable(child.table, context, references, `${path}[${index}].table`);
      }
      const name = Object.keys(child)[0] ?? "child";
      reject(
        "table",
        `${path}[${index}]`,
        name,
        "RTF table cells support paragraphs and nested tables only",
      );
    })
    .join("");
}

export function writeTable(
  table: TableOptions,
  context: RtfGenerateContext,
  references: NumberingReferences,
  path: string,
): string {
  assertAllowed(table, ["rows", "columnWidths"], "table", path);
  return table.rows
    .map((row, rowIndex) => {
      assertAllowed(row, ["cells"], "table", `${path}.rows[${rowIndex}]`);
      if (!("cells" in row) || row.cells.some((cell) => "sdt" in cell || "customXml" in cell)) {
        reject("table", `${path}.rows[${rowIndex}]`, "cells", "RTF rows support direct cells only");
      }
      const rowDefinition = (table.columnWidths ?? [])
        .map((width) => {
          if (typeof width !== "number") {
            reject("table", path, "columnWidths", "RTF column widths require numeric twips");
          }
          return control("cellx", width);
        })
        .join("");
      const cells = writeTableCells(
        row.cells as TableCellOptions[],
        context,
        references,
        `${path}.rows[${rowIndex}]`,
      );
      return `${control("trowd")}${rowDefinition}${cells}${control("row")}`;
    })
    .join("");
}
