import { describe, expect, it } from "vite-plus/test";

import { generateDocumentFile, parseDocumentFile } from "./document";
import { detectOfficeFormat } from "./formats";
import { generateOfficeDocument, generate } from "./generate";
import * as officeOpen from "./index";
import { generatePresentationFile, parsePresentationFile } from "./presentation";
import { SCHEMAS } from "./schemas";
import { generateWorkbookFile, parseWorkbookFile } from "./workbook";

const documentOptions = { sections: [{ children: [{ paragraph: "Family" }] }] };
const workbookOptions = { worksheets: [{ rows: [{ cells: [{ value: "Family" }] }] }] };
const presentationOptions = {
  slides: [
    {
      children: [{ shape: { x: 0, y: 0, width: 10, height: 4, textBody: { text: "Family" } } }],
    },
  ],
};

describe("family file APIs", () => {
  it("generates and parses each OpenDocument family", async () => {
    const document = await parseDocumentFile(
      await generateDocumentFile("odt", documentOptions, "uint8array"),
    );
    const workbook = await parseWorkbookFile(
      await generateWorkbookFile("ods", workbookOptions, "uint8array"),
    );
    const presentation = await parsePresentationFile(
      await generatePresentationFile("odp", presentationOptions, "uint8array"),
    );

    expect(document.sections).toHaveLength(1);
    expect(workbook.worksheets).toHaveLength(1);
    expect(presentation.slides).toHaveLength(1);
  });

  it("dispatches canonical formats through the root generator", async () => {
    const bytes = (await generateOfficeDocument(
      "odt",
      documentOptions,
      "uint8array",
    )) as Uint8Array;
    expect(detectOfficeFormat(bytes).format).toBe("odt");

    const legacy = (await generate({
      type: "ods",
      options: workbookOptions,
      outputType: "uint8array",
    })) as Uint8Array;
    expect(detectOfficeFormat(legacy).format).toBe("ods");
  });

  it("detects every supported OpenDocument generation format", () => {
    expect(detectOfficeFormat("{\\rtf1\\ansi Family}").family).toBe("document");
    expect(() => detectOfficeFormat(new TextEncoder().encode("not office"))).toThrow(
      "Unable to detect",
    );
  });

  it("exposes only the aggregate public runtime API", () => {
    expect(Object.keys(officeOpen).sort()).toEqual([
      "detectOfficeFile",
      "detectOfficeFormat",
      "generate",
      "generateDocumentFile",
      "generateOfficeDocument",
      "generatePresentationFile",
      "generateToFile",
      "generateWorkbookFile",
      "parseDocumentFile",
      "parseInput",
      "parseOfficeDocument",
      "parsePresentationFile",
      "parseWorkbookFile",
    ]);
  });

  it("keeps canonical ODF inputs on the three OOXML schemas", () => {
    expect(Object.keys(SCHEMAS).sort()).toEqual(["docx", "pptx", "xlsx"]);
  });
});
