import type {
  BorderOptions,
  SectionChild,
  ShadingProperties,
  TableOptions,
  TableCellOptions,
  TableRowOptions,
} from "@office-open/docx";

import { assertAllowed, registerColor, reject, type RtfGenerateContext } from "./context";
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
      assertAllowed(
        cell,
        [
          "children",
          "columnSpan",
          "rowSpan",
          "verticalAlign",
          "shading",
          "borders",
          "horizontalMerge",
          "verticalMerge",
        ],
        "table",
        `${path}.cells[${index}]`,
      );
      return (
        writeCellChildren(cell.children, context, references, `${path}.cells[${index}].children`) +
        control("cell")
      );
    })
    .join("");
}

function cellBorder(
  side: "top" | "bottom" | "left" | "right",
  border: BorderOptions,
  context: RtfGenerateContext,
  path: string,
): string {
  if (border.style !== "single")
    reject("table", path, border.style, "RTF table cell borders support single borders");
  return `${control(`clbrdr${side[0]}`)}${control("brdrs")}${
    border.size === undefined ? "" : control("brdrw", border.size)
  }${border.color === undefined ? "" : control("brdrcf", registerColor(context, border.color))}`;
}

function cellProperties(cell: TableCellOptions, context: RtfGenerateContext, path: string): string {
  let output = "";
  if (cell.horizontalMerge === "restart") output += control("clmgf");
  if (cell.horizontalMerge === "continue") output += control("clmrg");
  if (cell.verticalMerge === "restart") output += control("clvmgf");
  if (cell.verticalMerge === "continue") output += control("clvmrg");
  if (cell.verticalAlign === "top") output += control("clvertalt");
  if (cell.verticalAlign === "center") output += control("clvertalc");
  if (cell.verticalAlign === "bottom") output += control("clvertalb");
  const shading = cell.shading as ShadingProperties | undefined;
  if (shading) {
    if (shading.type !== "clear" || typeof shading.fill !== "string")
      reject("table", `${path}.shading`, "value", "RTF table cells support clear string fills");
    output += control("clcbpat", registerColor(context, shading.fill));
  }
  for (const side of ["top", "bottom", "left", "right"] as const) {
    const border = cell.borders?.[side];
    if (border) output += cellBorder(side, border, context, `${path}.borders.${side}`);
  }
  return output;
}

function rowProperties(row: TableRowOptions, path: string): string {
  let output = "";
  if (row.tableHeader) output += control("trhdr");
  if (row.cantSplit) output += control("trkeep");
  if (row.height) {
    if (typeof row.height.value !== "number" || row.height.rule !== "atLeast")
      reject("table", `${path}.height`, "value", "RTF row heights require atLeast twips");
    output += control("trh", row.height.value);
  }
  if (row.rowAlignment === "left") output += control("trql");
  if (row.rowAlignment === "center") output += control("trqc");
  if (row.rowAlignment === "right") output += control("trqr");
  return output;
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
      assertAllowed(
        row,
        ["cells", "cantSplit", "tableHeader", "height", "rowAlignment"],
        "table",
        `${path}.rows[${rowIndex}]`,
      );
      if (!("cells" in row) || row.cells.some((cell) => "sdt" in cell || "customXml" in cell)) {
        reject("table", `${path}.rows[${rowIndex}]`, "cells", "RTF rows support direct cells only");
      }
      let boundary = 0;
      const rowDefinition = (row.cells as TableCellOptions[])
        .map((cell, cellIndex) => {
          boundary += cell.columnSpan ?? 1;
          const width = table.columnWidths?.[boundary - 1];
          if (typeof width !== "number") {
            reject("table", path, "columnWidths", "RTF column widths require numeric twips");
          }
          return `${cellProperties(cell, context, `${path}.rows[${rowIndex}].cells[${cellIndex}]`)}${control("cellx", width)}`;
        })
        .join("");
      const rowControls = rowProperties(row as TableRowOptions, `${path}.rows[${rowIndex}]`);
      const cells = writeTableCells(
        row.cells as TableCellOptions[],
        context,
        references,
        `${path}.rows[${rowIndex}]`,
      );
      return `${control("trowd")}${rowControls}${rowDefinition}${cells}${control("row")}`;
    })
    .join("");
}
