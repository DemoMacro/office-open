import type {
  SectionChild,
  TableCellOptions,
  TableOptions,
  TableRowOptions,
} from "@office-open/docx";

import { DocParseError } from "../errors";
import type { LegacyParagraphProperties, ParagraphRange, TextCharacter } from "./models";
export function paragraphPropertiesFor(
  ranges: readonly ParagraphRange[],
  cp: number | undefined,
): LegacyParagraphProperties {
  if (cp === undefined) return {};
  const range = ranges.find((candidate) => candidate.cpStart <= cp && cp < candidate.cpEnd);
  if (!range) return {};
  const {
    alignment,
    indentLeft,
    indentRight,
    firstLine,
    spacingBefore,
    spacingAfter,
    spacingLine,
    outlineLevel,
  } = range.properties;
  const properties: LegacyParagraphProperties = {};
  if (alignment !== undefined) properties.alignment = alignment;
  if (indentLeft !== undefined || indentRight !== undefined || firstLine !== undefined) {
    properties.indent = {
      ...(indentLeft !== undefined ? { left: indentLeft } : {}),
      ...(indentRight !== undefined ? { right: indentRight } : {}),
      ...(firstLine !== undefined ? { firstLine } : {}),
    };
  }
  if (spacingBefore !== undefined || spacingAfter !== undefined || spacingLine !== undefined) {
    properties.spacing = {
      ...(spacingBefore !== undefined ? { before: spacingBefore } : {}),
      ...(spacingAfter !== undefined ? { after: spacingAfter } : {}),
      ...(spacingLine !== undefined ? { line: spacingLine, lineRule: "auto" as const } : {}),
    };
  }
  if (outlineLevel !== undefined) properties.outlineLevel = outlineLevel;
  return properties;
}

export function applyTables(
  children: SectionChild[],
  paragraphRanges: readonly ParagraphRange[],
  characters: readonly TextCharacter[],
): SectionChild[] {
  const rowByStart = new Map<number, ParagraphRange>();
  for (const range of paragraphRanges) {
    if (
      !range.properties.inTable ||
      !range.properties.rowEnd ||
      !range.properties.columnWidths?.length
    )
      continue;
    const existing = rowByStart.get(range.cpStart);
    rowByStart.set(range.cpStart, !existing || range.cpEnd > existing.cpEnd ? range : existing);
  }
  const rows: ParagraphRange[] = [];
  for (const range of [...rowByStart.values()].sort(
    (left, right) => left.cpStart - right.cpStart,
  )) {
    const previous = rows.at(-1);
    if (previous && range.cpStart < previous.cpEnd) {
      previous.cpEnd = Math.max(previous.cpEnd, range.cpEnd);
      previous.properties = { ...previous.properties, ...range.properties };
    } else {
      rows.push(range);
    }
  }
  if (rows.length === 0) return children;
  const result = [...children];
  let offset = 0;
  let previousEnd = 0;
  for (const row of rows) {
    const startIndex = paragraphIndexForCp(characters, previousEnd);
    const endIndex = paragraphIndexForCp(characters, row.cpEnd);
    const content = result.slice(startIndex + offset, endIndex + offset);
    const cellEnds = characters
      .filter(
        (character) => character.cellEnd && character.cp >= previousEnd && character.cp < row.cpEnd,
      )
      .map((character) => paragraphIndexForCp(characters, character.cp) + 1 - startIndex);
    const columnCount = row.properties.columnWidths!.length;
    if (cellEnds.length !== columnCount || cellEnds.at(-1) !== content.length) {
      throw new DocParseError("Invalid DOC table: cell markers do not match the row definition", {
        part: "document",
        path: "table/cell-boundaries",
        reason: "invalid-table-cell-boundaries",
      });
    }
    let cellStart = 0;
    const cells: TableCellOptions[] = cellEnds.map((cellEnd, index) => {
      const merge = row.properties.cellMerges?.[index];
      const cell = {
        children: content.slice(cellStart, cellEnd),
        ...(merge?.continue ? { verticalMerge: "continue" as const } : {}),
        ...(merge?.restart ? { verticalMerge: "restart" as const } : {}),
      };
      cellStart = cellEnd;
      return cell;
    });
    const table: SectionChild = {
      table: {
        columnWidths: row.properties.columnWidths,
        rows: [{ cells }] satisfies TableRowOptions[],
      } satisfies TableOptions,
    };
    result.splice(startIndex + offset, content.length, table);
    offset += 1 - content.length;
    previousEnd = row.cpEnd;
  }
  return result;
}

export function paragraphIndexForCp(characters: readonly TextCharacter[], cp: number): number {
  let count = 0;
  for (const character of characters) {
    if (character.paragraphEnd && character.cp < cp) count += 1;
  }
  return count;
}
