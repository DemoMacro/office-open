/**
 * Delimited-text conversion for RFC 4180 CSV/TSV and xlsx WorkbookOptions.
 *
 * Parsing reads CSV text as cell strings (no numeric/type guessing), strips a
 * leading UTF-8 BOM, accepts LF or CRLF record separators, preserves line
 * breaks inside quoted fields, and requires `""` for an embedded quote.
 * Serialization defaults to CRLF, escapes only fields containing the
 * delimiter/quote/line break, and can prepend a BOM.
 *
 * @module
 */
import type { CellOptions, WorkbookOptions, WorksheetOptions } from "@office-open/xlsx";
import { parseA1Cell } from "@office-open/xlsx";

const UTF8_BOM = "\uFEFF";
const CSV_DELIMITER = ",";
const TSV_DELIMITER = "\t";

/** Options for converting delimited text to WorkbookOptions. */
export interface CsvToWorkbookOptions {
  /** Field separator; exactly one character (default: "," for CSV). */
  delimiter?: string;
  /** Worksheet name for the imported rows (default: omitted). */
  sheetName?: string;
}

/** Options for converting a single-sheet WorkbookOptions to delimited text. */
export interface WorkbookToCsvOptions {
  /** Field separator; exactly one character (default: "," for CSV). */
  delimiter?: string;
  /** Record separator (default: CRLF per RFC 4180). */
  newline?: "\r\n" | "\n";
  /** Prepend a UTF-8 BOM (default: false). */
  bom?: boolean;
}

function resolveDelimiter(delimiter: string | undefined, fallback: string): string {
  const resolved = delimiter ?? fallback;
  if (resolved.length !== 1 || resolved === '"' || resolved === "\r" || resolved === "\n") {
    throw new Error("delimiter must be one character other than quote or line break");
  }
  return resolved;
}

function parseCsv(source: string, delimiter: string): string[][] {
  if (source.charCodeAt(0) === 0xfeff) source = source.slice(1);
  const records: string[][] = [];
  let fields: string[] = [];
  let field = "";
  let position = 0;
  let fieldStarted = false;

  const endField = () => {
    fields.push(field);
    field = "";
  };
  const endRecord = () => {
    endField();
    records.push(fields);
    fields = [];
    fieldStarted = false;
  };

  while (position < source.length) {
    const char = source[position]!;
    if (char === '"') {
      fieldStarted = true;
      if (field !== "") throw new Error("unexpected quote in unquoted field");
      position++;
      let closed = false;
      while (position < source.length) {
        const quoted = source[position]!;
        if (quoted !== '"') {
          field += quoted;
          position++;
          continue;
        }
        if (source[position + 1] === '"') {
          field += '"';
          position += 2;
          continue;
        }
        position++;
        closed = true;
        const next = source[position];
        if (next !== undefined && next !== delimiter && next !== "\r" && next !== "\n") {
          throw new Error("characters after closing quote");
        }
        break;
      }
      if (!closed) {
        throw new Error("unterminated quoted field");
      }
      continue;
    }
    if (char === delimiter) {
      fieldStarted = true;
      endField();
      position++;
      continue;
    }
    if (char === "\r" || char === "\n") {
      position += char === "\r" && source[position + 1] === "\n" ? 2 : 1;
      endRecord();
      continue;
    }
    field += char;
    fieldStarted = true;
    position++;
  }
  if (fieldStarted) endRecord();
  return records;
}

function stringifyCsv(records: string[][], delimiter: string, newline: "\r\n" | "\n"): string {
  return records
    .map((fields) =>
      fields
        .map((value) =>
          value.includes(delimiter) ||
          value.includes('"') ||
          value.includes("\r") ||
          value.includes("\n")
            ? `"${value.replaceAll('"', '""')}"`
            : value,
        )
        .join(delimiter),
    )
    .join(newline);
}

function recordsToWorksheet(records: string[][], sheetName?: string): WorksheetOptions {
  return {
    ...(sheetName ? { name: sheetName } : {}),
    rows: records.map((fields) => ({
      cells: fields.map((value): CellOptions => ({ value })),
    })),
  };
}

function singleWorksheet(workbook: WorkbookOptions): WorksheetOptions {
  const worksheets = workbook.worksheets ?? [];
  if (worksheets.length !== 1) {
    throw new Error("CSV conversion requires exactly one worksheet");
  }
  return worksheets[0]!;
}

function worksheetToRecords(worksheet: WorksheetOptions): string[][] {
  const grid: (CellOptions | undefined)[][] = [];
  let maxRow = 0;
  let maxColumn = 0;

  worksheet.rows?.forEach((row, rowIndex) => {
    row.cells?.forEach((cell, columnIndex) => {
      const reference = cell.reference ? parseA1Cell(cell.reference) : undefined;
      const rowNumber = reference?.row ?? row.rowNumber ?? rowIndex + 1;
      const columnNumber = reference?.col ?? columnIndex + 1;
      grid[rowNumber - 1] ??= [];
      grid[rowNumber - 1]![columnNumber - 1] = cell;
      maxRow = Math.max(maxRow, rowNumber);
      maxColumn = Math.max(maxColumn, columnNumber);
    });
  });

  const records: string[][] = [];
  for (let rowIndex = 0; rowIndex < maxRow; rowIndex++) {
    const fields: string[] = [];
    for (let columnIndex = 0; columnIndex < maxColumn; columnIndex++) {
      fields.push(stringifyCell(grid[rowIndex]?.[columnIndex]));
    }
    records.push(fields);
  }
  return records;
}

function stringifyCell(cell: CellOptions | undefined): string {
  if (!cell) return "";
  const value = cell.error ?? cell.value;
  if (value === undefined || value === null) return "";
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  if (value instanceof Date) return value.toISOString();
  if (value.text !== undefined) return value.text;
  return value.runs?.map((run) => run.text).join("") ?? "";
}

/** Convert RFC 4180 CSV text to WorkbookOptions. Cell values stay strings. */
export function csvToWorkbook(source: string, options: CsvToWorkbookOptions = {}): WorkbookOptions {
  const delimiter = resolveDelimiter(options.delimiter, CSV_DELIMITER);
  return {
    worksheets: [recordsToWorksheet(parseCsv(source, delimiter), options.sheetName)],
  };
}

/** Convert RFC 4180 TSV text to WorkbookOptions. Cell values stay strings. */
export function tsvToWorkbook(
  source: string,
  options: Omit<CsvToWorkbookOptions, "delimiter"> = {},
): WorkbookOptions {
  return csvToWorkbook(source, { ...options, delimiter: TSV_DELIMITER });
}

/** Convert a single-sheet WorkbookOptions to RFC 4180 CSV text. */
export function workbookToCsv(
  workbook: WorkbookOptions,
  options: WorkbookToCsvOptions = {},
): string {
  const delimiter = resolveDelimiter(options.delimiter, CSV_DELIMITER);
  return serializeRecords(workbook, delimiter, options.newline ?? "\r\n", options.bom ?? false);
}

/** Convert a single-sheet WorkbookOptions to RFC 4180 TSV text. */
export function workbookToTsv(
  workbook: WorkbookOptions,
  options: Omit<WorkbookToCsvOptions, "delimiter"> = {},
): string {
  return serializeRecords(workbook, TSV_DELIMITER, options.newline ?? "\r\n", options.bom ?? false);
}

function serializeRecords(
  workbook: WorkbookOptions,
  delimiter: string,
  newline: "\r\n" | "\n",
  bom: boolean,
): string {
  const csv = stringifyCsv(worksheetToRecords(singleWorksheet(workbook)), delimiter, newline);
  return bom ? UTF8_BOM + csv : csv;
}
