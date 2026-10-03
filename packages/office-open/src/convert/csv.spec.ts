import type { WorkbookOptions } from "@office-open/xlsx";
import { describe, expect, it } from "vitest";

import { csvToWorkbook, tsvToWorkbook, workbookToCsv, workbookToTsv } from "./csv";

describe("csvToWorkbook", () => {
  it("parses BOM, CRLF, quoted delimiters, escaped quotes, and embedded newlines", () => {
    const workbook = csvToWorkbook('\uFEFFname,note\r\n"a,""b""","line1\nline2"\r\n');
    expect(workbook.worksheets).toHaveLength(1);
    expect(workbook.worksheets![0]!.rows).toEqual([
      { cells: [{ value: "name" }, { value: "note" }] },
      { cells: [{ value: 'a,"b"' }, { value: "line1\nline2" }] },
    ]);
  });

  it("uses a custom delimiter and sheet name", () => {
    const workbook = csvToWorkbook("a;b", { delimiter: ";", sheetName: "Data" });
    expect(workbook.worksheets![0]!.name).toBe("Data");
    expect(workbook.worksheets![0]!.rows![0]!.cells).toEqual([{ value: "a" }, { value: "b" }]);
  });

  it("imports TSV and preserves empty fields", () => {
    const workbook = tsvToWorkbook("a\t\nc\td");
    expect(workbook.worksheets![0]!.rows).toEqual([
      { cells: [{ value: "a" }, { value: "" }] },
      { cells: [{ value: "c" }, { value: "d" }] },
    ]);
  });

  it("rejects invalid quoted CSV", () => {
    expect(() => csvToWorkbook('"unterminated')).toThrow("unterminated quoted field");
    expect(() => csvToWorkbook('"x"y')).toThrow("characters after closing quote");
    expect(() => csvToWorkbook('a"b')).toThrow("unexpected quote in unquoted field");
  });

  it("parses an empty quoted field as one empty cell", () => {
    expect(csvToWorkbook('""')).toEqual({
      worksheets: [{ rows: [{ cells: [{ value: "" }] }] }],
    });
    expect(csvToWorkbook("")).toEqual({ worksheets: [{ rows: [] }] });
    expect(csvToWorkbook("\uFEFF")).toEqual({ worksheets: [{ rows: [] }] });
  });
});

describe("workbookToCsv", () => {
  const workbook: WorkbookOptions = {
    worksheets: [
      {
        name: "Data",
        rows: [
          {
            cells: [
              { value: "a,b" },
              { value: 'say "hi"' },
              { value: "line1\nline2" },
              { value: 12 },
            ],
          },
        ],
      },
    ],
  };

  it("escapes fields and defaults to CRLF without a trailing record separator", () => {
    expect(workbookToCsv(workbook)).toBe('"a,b","say ""hi""","line1\nline2",12');
  });

  it("supports BOM, LF, custom delimiter, and TSV", () => {
    expect(workbookToCsv(workbook, { bom: true, newline: "\n" })).toBe(
      `\uFEFF"a,b","say ""hi""","line1\nline2",12`,
    );
    expect(workbookToCsv(workbook, { delimiter: ";" })).toBe('a,b;"say ""hi""";"line1\nline2";12');
    expect(workbookToTsv(workbook)).toBe('a,b\t"say ""hi"""\t"line1\nline2"\t12');
  });

  it("restores sparse references into their grid positions", () => {
    const sparse: WorkbookOptions = {
      worksheets: [
        {
          rows: [
            { rowNumber: 1, cells: [{ reference: "B1", value: "one" }] },
            { rowNumber: 3, cells: [{ reference: "A3", value: "three" }] },
          ],
        },
      ],
    };
    expect(workbookToCsv(sparse)).toBe(",one\r\n,\r\nthree,");
  });

  it("requires exactly one worksheet", () => {
    expect(() => workbookToCsv({ worksheets: [] })).toThrow("exactly one worksheet");
    expect(() =>
      workbookToCsv({
        worksheets: [{ name: "A" }, { name: "B" }],
      }),
    ).toThrow("exactly one worksheet");
  });
});

describe("CSV round trip", () => {
  it("preserves RFC 4180 text through WorkbookOptions", () => {
    const csv = 'a,"b""c","d\r\ne"';
    const workbook = csvToWorkbook(csv);
    expect(workbookToCsv(workbook)).toBe(csv);
  });
});
