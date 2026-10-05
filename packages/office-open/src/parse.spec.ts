import { OOXML_PACKAGE_FORMATS, type OoxmlPackageFormat } from "@office-open/core";
import { describe, expect, it } from "vite-plus/test";

import { generateOffice, type OfficeGenerateFormat, type OfficeOptionsFor } from "./generate";
import { parseOffice } from "./parse";

const OOXML_TYPES = Object.keys(OOXML_PACKAGE_FORMATS) as OoxmlPackageFormat[];

const OOXML_FIXTURES = {
  wordprocessing: { sections: [{ children: [{ paragraph: "Unified" }] }] },
  presentation: {
    slides: [
      {
        children: [{ shape: { x: 1, y: 1, width: 10, height: 4, textBody: { text: "Unified" } } }],
      },
    ],
  },
  spreadsheet: { worksheets: [{ rows: [{ cells: [{ value: "Unified" }] }] }] },
} as const;

describe("parseOffice", () => {
  it.for(OOXML_TYPES)("parses the %s package through the unified entry", async (type) => {
    const data = await generateOffice(
      type as OfficeGenerateFormat,
      OOXML_FIXTURES[OOXML_PACKAGE_FORMATS[type].family] as OfficeOptionsFor[OfficeGenerateFormat],
      "uint8array",
    );
    const parsed = await parseOffice(data as Uint8Array);
    expect(parsed.format).toBe(type);
    if (type === "docx") {
      if (parsed.format !== "docx") throw new Error("Expected DOCX");
      expect(parsed.options.sections).toHaveLength(1);
    }
    if (type === "xlsx") {
      if (parsed.format !== "xlsx") throw new Error("Expected XLSX");
      expect(parsed.options.worksheets).toHaveLength(1);
    }
    if (type === "pptx") {
      if (parsed.format !== "pptx") throw new Error("Expected PPTX");
      expect(parsed.options.slides).toHaveLength(1);
    }
  });

  it("parses RTF into document options", async () => {
    const parsed = await parseOffice("{\\rtf1\\ansi Unified}");
    if (parsed.format !== "rtf") throw new Error("Expected an RTF document");
    expect(parsed.options.sections).toHaveLength(1);
  });

  it("parses generated ODF packages through format detection", async () => {
    const document = await parseOffice(
      await generateOffice(
        "odt",
        { sections: [{ children: [{ paragraph: "ODT" }] }] },
        "uint8array",
      ),
    );
    const workbook = await parseOffice(
      await generateOffice(
        "ods",
        { worksheets: [{ rows: [{ cells: [{ value: "ODS" }] }] }] },
        "uint8array",
      ),
    );
    const presentation = await parseOffice(
      await generateOffice(
        "odp",
        {
          slides: [
            {
              children: [
                { shape: { x: 0, y: 0, width: 10, height: 4, textBody: { text: "ODP" } } },
              ],
            },
          ],
        },
        "uint8array",
      ),
    );
    expect([document.format, workbook.format, presentation.format]).toEqual(["odt", "ods", "odp"]);
  });
});
