import type { OoxmlPackageFormat } from "@office-open/core";
import { OOXML_PACKAGE_FORMATS } from "@office-open/core";
import type { DocumentOptions } from "@office-open/docx";
import { generateOdp } from "@office-open/odp";
import { generateOds } from "@office-open/ods";
import { generateOdt } from "@office-open/odt";
import { describe, expect, it } from "vite-plus/test";

import { generate, type GenerateOptionsMap, type GenerateType } from "./generate";
import { parseOfficeDocument } from "./parse";

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

describe("parseOfficeDocument", () => {
  it.for(OOXML_TYPES)("parses the %s package through the unified entry", async (type) => {
    const data = await generate({
      type,
      options: OOXML_FIXTURES[
        OOXML_PACKAGE_FORMATS[type].family
      ] as GenerateOptionsMap[GenerateType],
      outputType: "uint8array",
    });
    const parsed = await parseOfficeDocument(data as Uint8Array);
    expect(parsed.type).toBe(type);
    if (type === "docx") {
      if (parsed.type !== "docx") throw new Error("Expected DOCX");
      expect(parsed.options.sections).toHaveLength(1);
    }
    if (type === "xlsx") {
      if (parsed.type !== "xlsx") throw new Error("Expected XLSX");
      expect(parsed.options.worksheets).toHaveLength(1);
    }
    if (type === "pptx") {
      if (parsed.type !== "pptx") throw new Error("Expected PPTX");
      expect(parsed.options.slides).toHaveLength(1);
    }
  });

  it("parses RTF into document options", async () => {
    const parsed = await parseOfficeDocument("{\\rtf1\\ansi Unified}");
    if (parsed.type !== "rtf") throw new Error("Expected an RTF document");
    expect((parsed.options as DocumentOptions).sections).toHaveLength(1);
  });

  it("parses generated ODF packages through format detection", async () => {
    const document = await parseOfficeDocument(
      generateOdt({ sections: [{ children: [{ paragraph: "ODT" }] }] }),
    );
    const workbook = await parseOfficeDocument(
      generateOds({ worksheets: [{ rows: [{ cells: [{ value: "ODS" }] }] }] }),
    );
    const presentation = await parseOfficeDocument(
      generateOdp({
        slides: [
          {
            children: [{ shape: { x: 0, y: 0, width: 10, height: 4, textBody: { text: "ODP" } } }],
          },
        ],
      }),
    );
    expect([document.type, workbook.type, presentation.type]).toEqual(["odt", "ods", "odp"]);
  });
});
