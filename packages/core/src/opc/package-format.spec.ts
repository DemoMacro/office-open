import { describe, expect, it } from "vite-plus/test";

import { OOXML_PACKAGE_FORMATS, resolveMainPartContentType } from "./package-format";

const WORKBOOK_PATH = "xl/workbook.xml";
const DOCUMENT_PATH = "word/document.xml";
const SHEET_TYPE = OOXML_PACKAGE_FORMATS.xlsx.mainContentType;
const TEMPLATE_TYPE = OOXML_PACKAGE_FORMATS.xltx.mainContentType;
const LEGACY_TEMPLATE_TYPE = "application/vnd.ms-excel.template.main+xml";
const WORD_TYPE = OOXML_PACKAGE_FORMATS.docx.mainContentType;

const overrides = (partName: string, contentType: string) => ({
  overrides: [{ partName, contentType }],
});

describe("resolveMainPartContentType", () => {
  it("preserves the registered family main content type", () => {
    expect(
      resolveMainPartContentType(
        "spreadsheet",
        overrides(`/${WORKBOOK_PATH}`, TEMPLATE_TYPE),
        WORKBOOK_PATH,
        SHEET_TYPE,
      ),
    ).toBe(TEMPLATE_TYPE);
  });

  it("preserves the legacy spreadsheet template content type", () => {
    expect(
      resolveMainPartContentType(
        "spreadsheet",
        overrides(`/${WORKBOOK_PATH}`, LEGACY_TEMPLATE_TYPE),
        WORKBOOK_PATH,
        SHEET_TYPE,
      ),
    ).toBe(LEGACY_TEMPLATE_TYPE);
  });

  it("rejects a legacy alias outside its family", () => {
    expect(
      resolveMainPartContentType(
        "wordprocessing",
        overrides(`/${DOCUMENT_PATH}`, LEGACY_TEMPLATE_TYPE),
        DOCUMENT_PATH,
        WORD_TYPE,
      ),
    ).toBe(WORD_TYPE);
  });

  it("falls back for unknown declarations", () => {
    expect(
      resolveMainPartContentType(
        "spreadsheet",
        overrides(`/${WORKBOOK_PATH}`, "application/unknown+xml"),
        WORKBOOK_PATH,
        SHEET_TYPE,
      ),
    ).toBe(SHEET_TYPE);
  });

  it("falls back without a source table", () => {
    expect(resolveMainPartContentType("spreadsheet", undefined, WORKBOOK_PATH, SHEET_TYPE)).toBe(
      SHEET_TYPE,
    );
  });
});
