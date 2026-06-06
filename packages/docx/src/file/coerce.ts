/**
 * Coerce function for SectionChild type.
 *
 * Converts plain-object representations into their corresponding class
 * instances, enabling a JSON-friendly API.
 *
 * @module
 */
import { AltChunk } from "./alt-chunk/alt-chunk";
import { CustomXmlBlock, CustomXmlCell, CustomXmlRow } from "./custom-xml";
import type { FileChild } from "./file-child";
import { Paragraph } from "./paragraph/paragraph";
import {
  StructuredDocumentTagBlock,
  StructuredDocumentTagCell,
  StructuredDocumentTagRow,
} from "./sdt";
import type { SectionChild } from "./section-child";
import { SubDoc } from "./sub-doc/sub-doc";
import { TableOfContents } from "./table-of-contents/table-of-contents";
import { Table } from "./table/table";
import type { TableOptions } from "./table/table";
import { TableCell } from "./table/table-cell";
import type { TableCellOptions } from "./table/table-cell";
import { TableRow } from "./table/table-row";
import type { TableRowOptions } from "./table/table-row";
import { Textbox } from "./textbox/textbox";
import { BaseXmlComponent } from "./xml-components";

/**
 * Coerces a `SectionChild` (plain object or class instance) into a `FileChild`.
 *
 * - Class instances pass through unchanged.
 * - Plain objects are mapped to their corresponding constructor.
 */
export function coerceSectionChild(child: SectionChild): FileChild {
  if (child instanceof BaseXmlComponent) return child as FileChild;
  if ("paragraph" in child) return new Paragraph(child.paragraph);
  if ("table" in child) return new Table(preCoerceTableOptions(child.table));
  if ("toc" in child) {
    const { alias, ...options } = child.toc;
    return new TableOfContents(alias, options);
  }
  if ("textbox" in child) {
    const { children, ...rest } = child.textbox;
    return new Textbox({ ...rest, children: children?.map(coerceSectionChild) });
  }
  if ("sdt" in child) {
    const { properties, children } = child.sdt;
    return new StructuredDocumentTagBlock({
      properties,
      children: children?.map(coerceSectionChild),
    });
  }
  if ("altChunk" in child) return new AltChunk(child.altChunk);
  if ("subDoc" in child) return new SubDoc(child.subDoc);
  if ("customXml" in child) return new CustomXmlBlock(child.customXml);
  throw new Error("Unknown section child type");
}

/**
 * Recursively pre-coerce a TableOptions tree so that all nested children
 * are already class instances before reaching Table/TableRow/TableCell constructors.
 */
function preCoerceTableOptions(opts: TableOptions): TableOptions {
  return { ...opts, rows: opts.rows.map(preCoerceTableRow) };
}

function preCoerceTableRow(
  row: TableRow | StructuredDocumentTagRow | CustomXmlRow | TableRowOptions,
): TableRow | StructuredDocumentTagRow | CustomXmlRow | TableRowOptions {
  if (
    row instanceof TableRow ||
    row instanceof StructuredDocumentTagRow ||
    row instanceof CustomXmlRow
  )
    return row;
  return { ...row, cells: row.cells.map(preCoerceTableCell) };
}

function preCoerceTableCell(
  cell:
    | TableCell
    | StructuredDocumentTagCell
    | StructuredDocumentTagRow
    | CustomXmlCell
    | TableCellOptions,
):
  | TableCell
  | StructuredDocumentTagCell
  | StructuredDocumentTagRow
  | CustomXmlCell
  | TableCellOptions {
  if (
    cell instanceof TableCell ||
    cell instanceof StructuredDocumentTagCell ||
    cell instanceof StructuredDocumentTagRow ||
    cell instanceof CustomXmlCell
  )
    return cell;
  return {
    ...cell,
    children: cell.children.map((c) => (c instanceof BaseXmlComponent ? c : coerceSectionChild(c))),
  };
}
